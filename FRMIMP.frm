VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form FRMIMP 
   Caption         =   "Impresiones ..."
   ClientHeight    =   2505
   ClientLeft      =   2085
   ClientTop       =   2550
   ClientWidth     =   5190
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5190
   Begin ComctlLib.ProgressBar ProgBar 
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      _Version        =   327680
      Appearance      =   0
      MouseIcon       =   "FRMIMP.frx":0000
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reporte :"
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   4935
      Begin VB.TextBox txtnumfac2 
         Height          =   285
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtnumfac1 
         Height          =   285
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtserie 
         Height          =   285
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox CMBFGB 
         Height          =   315
         ItemData        =   "FRMIMP.frx":001C
         Left            =   1080
         List            =   "FRMIMP.frx":0029
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox rLista 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txtfecha 
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
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtnucleo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   0
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblfgb 
         Caption         =   "Nro. del                               al"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblfgb 
         Caption         =   "Serie :"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblfgb 
         Caption         =   "Documento:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label LblMensaje 
         Caption         =   "Ing. Nucleo :"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label lblreporte 
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
         Left            =   1080
         TabIndex        =   7
         Top             =   120
         Width           =   180
      End
   End
   Begin VB.CommandButton cerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   4680
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
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
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblProceso 
      Alignment       =   2  'Center
      Caption         =   "Procesando ..."
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   3375
   End
End
Attribute VB_Name = "FRMIMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim xl As Object
Dim Wfile As String
Dim WFORM As String
Dim LOC_NUCLEO As String * 1
Dim WDISPO As String * 1
Dim PSLOT As rdoQuery
Dim rlot_mayor As rdoResultset
Dim PROTER_SALDO As Currency
Dim LOC_FECHA_INC As String
Dim PSCAA As rdoQuery
Dim CAA_REPO  As rdoResultset
Dim WS_SALDO_ULTIMO As Currency
Dim LOC_CANCELA As Integer



Private Sub cerrar_Click()
On Error GoTo sal
If Trim(UCase(Wfile)) = "CLI_IMP" Then
  xl.ActiveWorkbook.Close False
End If
Unload FRMIMP

Exit Sub
sal:
'MsgBox Err.Number
If Err.Number = 91 Then
  Resume Next
End If
MsgBox Err.Number & " " & Err.Description, 48, Pub_Titulo
End Sub

Private Sub CMBFGB_Click()
txtSerie.SetFocus
End Sub

Private Sub CMBFGB_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  txtSerie.SetFocus
End If
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
LOC_CANCELA = 0
Dim wODBC As String
Reportes.Connect = PUB_ODBC
WDISPO = ""
SQ_OPER = 1
LEER_TRA_LLAVE
If tra_llave.EOF Then
   Screen.MousePointer = 0
   FRMIMP.Hide
   Exit Sub
End If
Screen.MousePointer = 0
Wfile = Trim(tra_llave(3))
WFORM = Trim(tra_llave(7))
lblreporte.Caption = Trim(tra_llave(1))
If Trim(UCase(Wfile)) = "LOTE_IMP" Then
  lblmensaje.Visible = True
  lblmensaje.Caption = "N u c l e o  : "
  rLista.AddItem "A  = Cortijo "
  rLista.AddItem "B  = ()"
  rLista.AddItem "G  = ()"
  rLista.AddItem "M  = Milagro"
  rLista.AddItem "S  = Santo D."
  rLista.AddItem "R  = Reproductora"
  rLista.ListIndex = 0
  rLista.Visible = True
 ' LblMensaje.Visible = True
'  txtnucleo.Visible = True
  Exit Sub
Else
  lblmensaje.Visible = False
  txtNucleo.Visible = False
End If
If Trim(UCase(Wfile)) = "VENTAS_IMP" Then
  lblmensaje.Visible = True
  lblmensaje.Caption = "Fecha de Inicio: dd /mm /yyyy  ="
  txtfecha.Visible = True
  txtfecha.Text = Format(DateAdd("d", -19, LK_FECHA_DIA), "dd/mm/yyyy")
ElseIf Trim(UCase(Wfile)) = "VENTASMES_IMP" Then
  lblmensaje.Visible = True
  lblmensaje.Caption = "Fecha de Inicio: dd /mm /yyyy  ="
  txtfecha.Visible = True
  txtfecha.Text = Format(DateAdd("d", -30, LK_FECHA_DIA), "dd/mm/yyyy")
ElseIf Trim(UCase(Wfile)) = "CLI_IMP" Or Trim(UCase(Wfile)) = "VENDIA.RPT" Or Trim(UCase(Wfile)) = "ABONOS_IMP" Or Trim(UCase(Wfile)) = "CLI_SALDOS" Then
  lblmensaje.Visible = True
  lblmensaje.Caption = "Fecha de Inicio: dd /mm /yyyy  ="
  txtfecha.Visible = True
  txtfecha.Text = Format(DateAdd("d", -12, LK_FECHA_DIA), "dd/mm/yyyy")
ElseIf Trim(UCase(Wfile)) = "FORMULA_IMP" Then
  lblmensaje.Visible = True
  lblmensaje.Caption = "Clave Plantillla ="
  rLista.AddItem "Pollos - Carne " & String(60, " ") & "B"
  rLista.AddItem "Reproductora COBB" & String(60, " ") & "R"
  rLista.AddItem " Reproductora ROSS" & String(60, " ") & "R.R"
  rLista.AddItem " Pollo-Carne ROSS MACHOS" & String(60, " ") & "B.R"
  rLista.AddItem " Pollo-Carne ROSS HEMBRAS " & String(60, " ") & "B.H"
  rLista.ListIndex = 0
  rLista.Visible = True
ElseIf Trim(UCase(Wfile)) = "IMP_FACTURACION" Then
  lblfgb(0).Visible = True
  lblfgb(1).Visible = True
  lblfgb(2).Visible = True
  lblmensaje.Caption = ""
  CMBFGB.Visible = True
  txtSerie.Visible = True
  txtnumfac1.Visible = True
  txtnumfac2.Visible = True
  txtfecha.Visible = False
Else
  txtNucleo.Visible = True
  lblmensaje.Visible = False
  txtfecha.Visible = False
End If


If Trim(UCase(Wfile)) = "INSUMOS.RPT" Then
  pub_cadena = "SELECT * FROM ARTICULO WHERE ARM_CODCIA = ? AND ARM_CODART = ? "
  Set PSLOT = CN.CreateQuery("", pub_cadena)
  Set rlot_mayor = PSLOT.OpenResultset(rdOpenKeyset, rdConcurValues)


PUB_KEY = 0
pu_codcia = LK_CODCIA
SQ_OPER = 2
LEER_ART_LLAVE
PROTER_SALDO = 0
Do Until art_mayor.EOF
   If art_mayor!art_familia = 12 Then
     PSLOT(0) = LK_CODCIA
     PSLOT(1) = art_mayor!ART_KEY
     rlot_mayor.Requery
     PROTER_SALDO = PROTER_SALDO + rlot_mayor!arm_stock
   End If
art_mayor.MoveNext
Loop
End If
'Unload FORMGEN
End Sub

Private Sub Pantalla_Click()
If Pantalla.Caption = "Cancelar" Then
  LOC_CANCELA = 1
  DoEvents
  Exit Sub
End If
PRO_REPORTE "P"
End Sub

Private Sub txtfecha_GotFocus()
Azul txtfecha, txtfecha
End Sub

Private Sub txtfecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Pantalla.Enabled Then
    Pantalla.SetFocus
  End If
End If
End Sub

Private Sub txtnucleo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Pantalla.Enabled Then
    Pantalla.SetFocus
  End If
End If
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

unpoco = 0
cad = "SELECT * FROM COMAEST WHERE COM_CUENTA >= ? and COM_NIVEL <> 2  and COM_CODCIA = ? ORDER BY COM_CUENTA"
Set PC_CUENTA = CN.CreateQuery("", cad)
Set ps_cta = PC_CUENTA.OpenResultset(rdOpenKeyset, rdConcurValues)

WTABULA2 = String(1, " ")
WTABULA = String(1, " ")
RUTA = PUB_RUTA_REPORTE & Trim(LK_CODUSU) & ".txt"
Close #1
If LKCHEK Then Open RUTA For Output As #1


SQ_OPER = 2
PUB_TIPREG = WT_ESTADO
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
wSUMGRUPO1 = 0
BAN_GRUPO = "N"
una_ves = ""
sp_grupo = 0
WR_PAG = 0
If LKCHEK = False Then GoSub Wtamano

GoSub CABEZA
Do Until tab_mayor.EOF ' LOOP 1
  PW_VALOR1 = Trim(tab_mayor!tab_nomlargo)
  PW_VALOR2 = Trim(tab_mayor!tab_nomcorto)
  CTA_SIGNO = Val(tab_mayor!TAB_CODART)
  CTA_RESTA = Trim(tab_mayor!TAB_contable2)
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
    wp_SUMTOTAL2 = ALIM(wp_SUMTOTAL2)
    wp_DESCRIPCION = PW_CUENTA
    If wSUMTOTAL2 < 0 Then
      unpoco = 1
    Else
      unpoco = 0
    End If
    If Not LKCHEK Then Printer.FontBold = True
    If LKCHEK Then Print #1, WTABULA2; Spc(62); " ---------" Else Printer.Print WTABULA2; Spc(62); " ---------"
    If LKCHEK Then Print #1, WTABULA2; Spc(30); wp_DESCRIPCION; "   "; Spc(unpoco); wp_SUMTOTAL2 _
    Else Printer.Print WTABULA2; Spc(30); wp_DESCRIPCION; "   "; Spc(unpoco); wp_SUMTOTAL2
    If LKCHEK Then Print #1, WTABULA2; "" Else Printer.Print WTABULA2; ""
    If Not LKCHEK Then Printer.FontBold = False
    'wSUMTOTAL2 = 0
    GoTo OTRO
  End If
  If PW_VALOR2 = "T" Then
    If LKCHEK Then Print #1, WTABULA2; Spc(2); PW_CUENTA Else Printer.Print WTABULA2; Spc(2); PW_CUENTA
    If LKCHEK Then Print #1, WTABULA2; "" Else Printer.Print WTABULA2; ""
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
    'If LKCHEK Then Print #1, WTABULA2; Spc(52); " ---------" Else Printer.Print WTABULA2; Spc(52); " ---------"
    If LKCHEK Then Print #1, WTABULA2; Spc(28); wp_DESCRIPCION; "     "; Spc(unpoco); wp_SUMGRUPO1 Else Printer.Print WTABULA2; Spc(28); wp_DESCRIPCION; "     "; Spc(unpoco); wp_SUMGRUPO1
    If LKCHEK Then Print #1, WTABULA2; "" Else Printer.Print WTABULA2; ""
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
  NIVELES = ps_cta!COM_NIVEL
  un_nivel = 0
  Do Until ps_cta.EOF
    NIVELES = ps_cta!COM_NIVEL
    If ps_cta!COM_NIVEL = "1" Then
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
          CTA_RESTA_SOLES = ((com_llave!COM_DEB_ANO + com_llave!COM_DEB_MES) * com_llave!com_SIGNO_D) + ((com_llave!COM_HAB_ANO + com_llave!COM_HAB_MES) * com_llave!com_SIGNO_H)
          If CTA_RESTA_SOLES <> 0 Then CTA_RESTA_SOLES = CTA_RESTA_SOLES * -1
      End If
     End If
    If CTA_SIGNO = 0 Then
      WMONTO = ((Val(ps_cta!COM_HAB_ANO) + Val(ps_cta!COM_HAB_MES)) * ps_cta!com_SIGNO_H) + ((Val(ps_cta!COM_DEB_ANO) + Val(ps_cta!COM_DEB_MES)) * ps_cta!com_SIGNO_D) + CTA_RESTA_SOLES
    Else
      WMONTO = ((Val(ps_cta!COM_HAB_ANO) + Val(ps_cta!COM_HAB_MES)) * ps_cta!com_SIGNO_H) + ((Val(ps_cta!COM_DEB_ANO) + Val(ps_cta!COM_DEB_MES)) * ps_cta!com_SIGNO_D) * CTA_SIGNO + CTA_RESTA_SOLES
    End If
    If ps_cta!COM_NIVEL = 1 Then
       wSUMTOTAL1 = wSUMTOTAL1 + WMONTO
       wSUMTOTAL2 = wSUMTOTAL2 + WMONTO
    End If
    wp_DESCRIPCION = ps_cta!com_DESCRIPCION
    wp_MONTO = Format(WMONTO, "##,###,###.00")
    If Not LKCHEK Then Printer.FontBold = False
    If ps_cta!COM_NIVEL = 1 Then
      If Not LKCHEK Then Printer.FontBold = True
      wp_MONTO = NUM_NEGA(WMONTO)
      wp_DESCRIPCION = UCase(ps_cta!com_DESCRIPCION)
      spacio = 0
      If BAN_GRUPO = "S" Then
       spacio2 = 8 + 6
      Else
       spacio2 = 22
      End If
    ElseIf ps_cta!COM_NIVEL = 2 Then
      If WMONTO < 0 Then
        wp_MONTO = Format(WMONTO * -1, "##,###,###.00")
      End If
      spacio = 5
      spacio2 = 8
    ElseIf ps_cta!COM_NIVEL = 3 Then
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
    wp_MONTO = BAN_LINE(wp_MONTO)
    If WMONTO < 0 Then
      unpoco = 1
    Else
     unpoco = 0
    End If
    If WMONTO <> 0 Then
       If LKCHEK Then Print #1, WTABULA2; Spc(spacio); wp_CUENTA; Spc(1); wp_DESCRIPCION; Spc(spacio2 + sp_grupo + unpoco); wp_MONTO _
       Else Printer.Print WTABULA2; Spc(spacio); wp_CUENTA; Spc(1); wp_DESCRIPCION; Spc(spacio2 + sp_grupo + unpoco); wp_MONTO
    End If
    If Not LKCHEK Then Printer.FontBold = False
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
Close #1
If LKCHEK Then DoEvents: frmeditor.Show 1 Else Printer.EndDoc




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

CABEZA:
    If LKCHEK Then Print #1, " " Else Printer.Print " "
    WR_PAG = WR_PAG + 1
    WR_FECHA = Format(LK_FECHA_DIA, "dddd, dd mmmm yyyy")
    WR_CIA = Mid(MDIForm1.StatusBar1.Panels(7).Text, 4, Len(MDIForm1.StatusBar1.Panels(7).Text) - 3)
    If LKCHEK Then Print #1, Spc(10); WR_CIA Else Printer.Print Spc(10); WR_CIA
    If LKCHEK Then Print #1, Spc(10); String(15, "-") Else Printer.Print Spc(10); String(15, "-")
    If LKCHEK Then Print #1, " " Else Printer.Print " "
    If WT_ESTADO = 77 Then
       If LKCHEK Then Print #1, Spc(15); " EST. G. P. POR FUNCION           "; Spc(15); "Pag. : "; WR_PAG _
        Else: Printer.Print Spc(15); " EST. G. P. POR FUNCION           "; Spc(15); "Pag. : "; WR_PAG
    Else
       If LKCHEK Then Print #1, Spc(15); " EST. G. P. POR NETURALEZA        "; Spc(15); "Pag. : "; WR_PAG _
        Else: Printer.Print Spc(15); " EST. G. P. POR NETURALEZA        "; Spc(15); "Pag. : "; WR_PAG
    End If
    If LKCHEK Then Print #1, Spc(15); "DEL 01 AL "; Spc(1); WR_FECHA _
    Else: Printer.Print Spc(15); ; "DEL 01 AL "; Spc(1); WR_FECHA
    If LKCHEK Then Print #1, Spc(15); String(37, "-") Else Printer.Print Spc(1)
    If LKCHEK Then Print #1, " " Else Printer.Print " "
    If LKCHEK Then Print #1, " " Else Printer.Print " "
    If LKCHEK Then Print #1, " " Else Printer.Print " "
    
Return

Wtamano:
Printer.Print " "
Printer.FontSize = 11#     '7
Printer.FontName = "Courier New"
Printer.FontBold = False
Return


fin:
End Sub

Public Function ALIM(VAR As String) As String
Dim TEMP As String * 15
Dim N1 As Integer
Dim N2 As Integer
N1 = InStr(1, VAR, " ") - 1
N2 = Len(VAR) - N1
VAR = String(N2, "    ") + Left(VAR, N1) + String(1, " ")
ALIM = VAR
End Function

Public Sub CLI_SALDOS()
' *** REPORTES DE NUCLEOS
On Error GoTo FINTODO
Dim success%
Dim LETRAS(24) As String * 1
Dim MESES(12) As String
Dim FECHA_INICIO
Dim WSFECHA As Date
Dim xMES1, xMES2 As Integer
Dim con  As Integer
Dim TABLAGEN() As String * 30
Dim TABLAGEN2() As String * 30

Dim todos_cli As Integer
Dim filas_cli As Integer
Dim xpollos As String * 6
Dim xdias As String * 6
Dim C1 As Integer
Dim F1 As Integer
Dim WTABULA2 As String * 1
Dim WTABULA As String * 1
Dim RUTA As String
Dim WR_FECHA As String
Dim WR_LOTE As String
Dim WR_DAT1 As String
Dim SUM_cantidad As Currency
Dim WR_soles As String * 10
Dim WR_CIA As String
Dim cade As String

Dim SUM_SALDOS As Currency
Dim w_cant As String * 12
Dim w_sol As String * 10
Dim WCANT As Currency

Dim nuc_repo  As rdoResultset
Dim PSNUC_REPO As rdoQuery
Dim Mensaje, titulo, valorpred As String
Dim wvalor
Dim valor As Double
Dim xcuenta As Integer
Dim PAG_HOJAS As Integer
Dim WS_SALDO As Currency
Dim WS_S As String
Dim ws_tt As String

'FRMIMP.pantalla.Caption = "Cancelar"
'FRMIMP.pantalla.Enabled = True
DoEvents
cade = "SELECT * FROM CARACU WHERE CAA_CODCIA = ? AND CAA_CP = ? AND CAA_CODCLIE = ? AND CAA_FECHA >= ?  ORDER BY CAA_CODCLIE, CAA_FECHA, CAA_NUM_OPER"
Set PSCAA = CN.CreateQuery("", cade)
Set CAA_REPO = PSCAA.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

cade = "SELECT * FROM CLIENTES WHERE CLI_CODCIA = ? AND CLI_CP = ? AND CLI_ESTADO = 'B'  ORDER BY CLI_NOMBRE"
Set PSNUC_REPO = CN.CreateQuery("", cade)
Set nuc_repo = PSNUC_REPO.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

If wver = 99 Then FRMIMP.ProgBar.Value = 1


'*** LLENA DE TABLAS ... CLIENTES NUCLEOS
FECHA_INICIO = LOC_FECHA_INC
PSNUC_REPO(0) = LK_CODCIA
PSNUC_REPO(1) = "C"
Screen.MousePointer = 11
SUM_SALDOS = 0
nuc_repo.Requery
If nuc_repo.EOF Then
  Screen.MousePointer = 0
  MsgBox "NO EXISTE NINGUNO ..", 48, Pub_Titulo
  FRMIMP.ProgBar.Visible = False
  FRMIMP.lblproceso.Visible = False
  Exit Sub
End If
FRMIMP.ProgBar.Min = 0
FRMIMP.ProgBar.Max = nuc_repo.RowCount
FRMIMP.ProgBar.Value = 0
FRMIMP.ProgBar.Visible = True
DoEvents
FRMIMP.lblproceso.Visible = True
FRMIMP.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
todos_cli = nuc_repo.RowCount
'todos_cli = 26
ReDim TABLAGEN(32, todos_cli)
ReDim TABLAGEN2(32, todos_cli)
Dim wFECHASIGUE As Date
Dim wtC, wtF

GoSub WEXCEL
PAG_HOJAS = 1

xl.Worksheets(1).Activate
GoSub LETRAS
WS_SALDO = 0
xcuenta = 0
F1 = 1
Do Until nuc_repo.EOF ' loop 1
  xcuenta = xcuenta + 1
  If LOC_CANCELA = 1 Then
    GoTo CANCELA
  End If
  FRMIMP.ProgBar.Value = xcuenta
  C1 = 0
  TABLAGEN(C1, F1) = Trim(nuc_repo!cli_nombre)
  SUM_cantidad = 0
  DoEvents
  FRMIMP.lblproceso.Caption = Trim(nuc_repo!cli_nombre) & " . . .  " & xcuenta & " de " & todos_cli
  QSALDO = 0
  QSALDO1 = 0
  QJ = 0
  wdia = 0
  WSFECHA = CDate(FECHA_INICIO)
  wFECHASIGUE = WSFECHA
  FRMIMP.ProgBar.Min = 0
  FRMIMP.ProgBar.Max = 13
  FRMIMP.ProgBar.Value = 0
  Do Until wdia >= 13
    FRMIMP.ProgBar.Value = wdia
    wFECHASIGUE = DateAdd("d", wdia, WSFECHA)
    If CDate(wFECHASIGUE) > CDate(LK_FECHA_DIA) Then
      Exit Do
    End If
    C1 = C1 + 1
    ws_total = 0
    ws_tt = ""
    TABLAGEN(C1, 0) = CDate(wFECHASIGUE)
    FRMIMP.ProgBar.Value = FRMIMP.ProgBar.Max
    xl.Worksheets("Hoja1").Range(LETRAS(C1 + 2) & F1 + 7).HorizontalAlignment = xlRight
    wdia = wdia + 1
  Loop
    fx = 0
    PSCAA(0) = LK_CODCIA
    PSCAA(1) = "C"
    PSCAA(2) = nuc_repo!cli_codclie
    PSCAA(3) = CDate("7/11/98")
    CAA_REPO.Requery
    If CAA_REPO.EOF Then
       GoTo Siguef
    End If
    WDIF = DateDiff("d", WSFECHA, CAA_REPO!CAA_FECHA) + 3
    If WDIF > 3 Then
      QJ = 3
      LOC_VALOR = 0
      GoSub LLENA_VALOR
    End If
    QFECHA = CAA_REPO!CAA_FECHA
    QSALDO = CAA_REPO!CAA_SALDO
    QSALDO1 = 0
    QSALDO2 = 0
        
    Do Until CAA_REPO.EOF
    If QFECHA = CAA_REPO!CAA_FECHA Then
        GoTo SD
    Else
       WDIF = DateDiff("d", WSFECHA, CAA_REPO!CAA_FECHA) + 3
       
       If fx = 0 Then QSALDO = QSALDO2
       
       If WDIF >= 3 Then
         If QJ >= 3 Then xl.Cells(F1 + 7, QJ) = Format(QSALDO, "0.00")
               QJ = QJ + 1
               LOC_VALOR = QSALDO1
               GoSub LLENA_VALOR
       End If
        QFECHA = CAA_REPO!CAA_FECHA
        fx = 0
    End If
   QSALDO2 = QSALDO1
SD:
   
    QSALDO1 = CAA_REPO!CAA_SALDO

     If CAA_REPO!CAA_IMPORTE = CAA_REPO!caa_SALDO_car And CAA_REPO!CAA_IMPORTE > 0 Then
     Else
         fx = 1
         QSALDO = CAA_REPO!CAA_SALDO
     End If
     QJ = WDIF
     CAA_REPO.MoveNext
   Loop
   
   If fx = 0 Then QSALDO = QSALDO2
   WDIF = DateDiff("d", WSFECHA, LK_FECHA_DIA) + 3
   
   If WDIF >= QJ Then
   If WDIF >= 3 Then
        If QJ >= 3 Then xl.Cells(F1 + 7, QJ) = Format(QSALDO, "0.00")
           QJ = QJ + 1
           LOC_VALOR = QSALDO1
           GoSub LLENA_VALOR
  End If
 End If
Siguef:
 FRMIMP.ProgBar.Max = todos_cli
 nuc_repo.MoveNext
 F1 = F1 + 1
Loop ' loop 1
FRMIMP.ProgBar.Value = todos_cli
Screen.MousePointer = 0
GoTo VAMOS

WEXCEL:
  Dim xlchart As Chart
  Dim wranF, wran1, wran2, WPAS
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FRMIMP.lblproceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
  DoEvents
  WPAS = "131296"
  xl.Workbooks.Open "C:\ADMIN\OFFICE\SALDOS.xls", 0, True, 4, WPAS, WPAS
Return

VAMOS:
  DoEvents
  FRMIMP.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  DoEvents
  xcuenta = 1
  For F1 = 1 To todos_cli
    xl.Cells(F1 + 7, 2) = Trim(TABLAGEN(0, F1))
    xl.Cells(F1 + 7, 1) = F1
  Next F1
  xcuenta = 0
  For F1 = 1 To todos_cli
    If xcuenta <= (todos_cli) Then
      xl.Cells(xcuenta + 7, 17) = " "
    End If
    xcuenta = xcuenta + 2
  Next F1
  
  WSFECHA = CDate(FECHA_INICIO)
  For C1 = 1 To 14
    wFECHASIGUE = DateAdd("d", C1 - 1, WSFECHA)
    xl.Cells(6, C1 + 2) = wFECHASIGUE
    xl.Cells(7, C1 + 2) = wFECHASIGUE
  Next C1
  xl.Cells((todos_cli) + 8, 2) = " TOTALES "
 
  wranF = "A" & (todos_cli) + 8 & ":B" & (todos_cli) + 9
  xl.Range(wranF).Font.Bold = True
 
  'wranF = "S8:S" & (todos_cli * 2) + 9
  'xl.Range(wranF).Font.Bold = True
  
  For F1 = 3 To 16
    wranF = LETRAS(F1) & (todos_cli) + 8
    wran1 = LETRAS(F1) & 8
    wran2 = LETRAS(F1) & (todos_cli) + 8 - 1
    xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  Next F1
  
  wran1 = "P8"
  wran2 = "P" & (todos_cli) + 7
  wranF = "P" & (todos_cli) + 8
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  
  wranF = "C" & (todos_cli) + 8 & ":P" & (todos_cli) + 8
  xl.Worksheets("Hoja1").Range(wranF).HorizontalAlignment = xlRight
  
  wranF = "A7:P" & (todos_cli) + 8
  xl.Range(wranF).Borders.LineStyle = 1

  WSFECHA = DateAdd("d", 15, CDate(FECHA_INICIO))
  xMES1 = Month(FECHA_INICIO)
  xMES2 = Month(WSFECHA)
  If xMES1 = xMES2 Then
    cade = Format(FECHA_INICIO, "mmmm")
  Else
    cade = Format(FECHA_INICIO, "mmmm") & " - " & Format(WSFECHA, "mmmm")
  End If
  xl.Cells(2, 2) = "R: " & PUB_CODTRA
  xl.Cells(3, 6) = "               S A L D O S  "
  xl.Worksheets("Hoja1").Range("A1:X51").Locked = True
  xl.Worksheets("Hoja1").Protect "131296"
  xl.Worksheets(1).Activate
  xl.Application.Visible = True
  DoEvents
  FRMIMP.lblproceso.Visible = False
  FRMIMP.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FRMIMP.Pantalla.Enabled = True
  FRMIMP.Pantalla.Caption = "Por &Pantalla"
  FRMIMP.lblproceso.Visible = False

Exit Sub



'*** RUTINAS PARA IMPRIMIR



WPROGRESO:

Return

Exit Sub
'SALE2:
'  MsgBox " Inesperado ..  " & Err.Description, 48, pub_titulo
 
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

LLENA_MESES:
MESES(1) = "ENERO"
MESES(1) = "FEBRERO"
MESES(1) = "MARZO"
MESES(1) = "ENERO"
MESES(1) = "ENERO"
MESES(1) = "ENERO"
MESES(1) = "ENERO"
MESES(1) = "ENERO"
MESES(1) = "ENERO"
MESES(1) = "ENERO"

Return

LLENA_VALOR:
For i = QJ To WDIF
  If i >= 3 Then
    xl.Cells(F1 + 7, i) = Format(LOC_VALOR, "0.00")
  End If
Next i
Return

Exit Sub
CANCELA:
  FRMIMP.Pantalla.Enabled = True
  FRMIMP.Pantalla.Caption = "Por &Pantalla"
  FRMIMP.lblproceso.Visible = False
  LOC_CANCELA = 0
  Set xl = Nothing
  Screen.MousePointer = 0

Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FRMIMP

End Sub


Public Sub CLI_COBRANZAS()
On ERRO GoTo FINTODO
Dim success%
Dim wranF, wran1, wran2, WPAS
Dim LETRAS(24) As String * 1
Dim MESES(12) As String
Dim FECHA_INICIO
Dim WSFECHA As Date
Dim xMES1, xMES2 As Integer
Dim con  As Integer
Dim TABLAGEN() As String * 30
Dim COL9, COL8, COL7, COL6, COL5, COL4, COL3, COL10
Dim TABLAGEN2() As String * 30
Dim i, WS_SUBGRUPO
Dim wt_kilos As Currency
Dim WS_PRECIO As Currency
Dim todos_cli As Integer
Dim filas_cli As Integer
Dim xpollos As String * 6
Dim xdias As String * 6
Dim C1 As Integer
Dim F1 As Integer
Dim WTABULA2 As String * 1
Dim WTABULA As String * 1
Dim RUTA As String
Dim WR_FECHA As String
Dim WR_LOTE As String
Dim WR_DAT1 As String
Dim SUM_cantidad As Currency
Dim WR_soles As String * 10
Dim WR_CIA As String
Dim cade As String
Dim tab_importe(10) As Currency
Dim TAB_CONCEPTO(10) As String * 30
Dim tab_kilos(10) As Currency
Dim tab_precio(10) As Currency

Dim SUM_SALDOS As Currency
Dim w_cant As String * 12
Dim w_sol As String * 10
Dim WCANT As Currency

Dim nuc_repo  As rdoResultset
Dim PSNUC_REPO As rdoQuery
Dim far_r  As rdoResultset
Dim PSFAR As rdoQuery

Dim Mensaje, titulo, valorpred As String
Dim wvalor
Dim valor As Double
Dim xcuenta As Integer
Dim PAG_HOJAS As Integer
Dim WS_SALDO As Currency
Dim WS_S As String
Dim ws_tt As String
Dim TOTAL_SALDO As Currency
'FRMIMP.pantalla.Caption = "Cancelar"
'FRMIMP.pantalla.Enabled = True
DoEvents
cade = "SELECT * FROM CARACU WHERE CAA_CODCIA = ? AND CAA_CP = ? AND CAA_CODCLIE = ? AND CAA_FECHA >= ?  ORDER BY CAA_CODCLIE, CAA_FECHA, CAA_NUM_OPER"
Set PSCAA = CN.CreateQuery("", cade)
Set CAA_REPO = PSCAA.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

cade = "SELECT * FROM CLIENTES WHERE CLI_CODCIA = ? AND CLI_CP = ? AND CLI_ESTADO = 'B'  ORDER BY CLI_SUBGRUPO, CLI_NOMBRE"
Set PSNUC_REPO = CN.CreateQuery("", cade)
Set nuc_repo = PSNUC_REPO.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

cade = "SELECT * FROM facart WHERE FAR_CODCIA = ? AND FAR_CP = ? AND FAR_CODCLIE = ? AND FAR_FECHA = ? AND FAR_TIPMOV = 28  AND FAR_ESTADO <> 'E' ORDER BY FAR_PRECIO"
Set PSFAR = CN.CreateQuery("", cade)
Set far_r = PSFAR.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

If wver = 99 Then FRMIMP.ProgBar.Value = 1
LOC_FECHA_INC = DateAdd("d", -1, LK_FECHA_DIA)

'*** LLENA DE TABLAS ... CLIENTES NUCLEOS
FECHA_INICIO = LOC_FECHA_INC
PSNUC_REPO(0) = LK_CODCIA
PSNUC_REPO(1) = "C"
Screen.MousePointer = 11
SUM_SALDOS = 0
nuc_repo.Requery
If nuc_repo.EOF Then
  Screen.MousePointer = 0
  MsgBox "NO EXISTE NINGUNO ..", 48, Pub_Titulo
  FRMIMP.ProgBar.Visible = False
  FRMIMP.lblproceso.Visible = False
  Exit Sub
End If
FRMIMP.ProgBar.Min = 0
FRMIMP.ProgBar.Max = nuc_repo.RowCount
FRMIMP.ProgBar.Value = 0
FRMIMP.ProgBar.Visible = True
DoEvents
FRMIMP.lblproceso.Visible = True
FRMIMP.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
todos_cli = nuc_repo.RowCount
Dim wFECHASIGUE As Date
Dim wtC, wtF

GoSub WEXCEL
PAG_HOJAS = 1

xl.Worksheets(1).Activate
GoSub LETRAS
WS_SALDO = 0
xcuenta = 0
F1 = 5
WS_SUBGRUPO = ""
FILA_INICIAL = 5
Do Until nuc_repo.EOF ' loop 1
    If nuc_repo!CLI_SUBGRUPO <> WS_SUBGRUPO Then
       If WS_SUBGRUPO <> "" Then GoSub TOTALES
       If nuc_repo!CLI_SUBGRUPO = 50 Then
          GoSub GENERAL
       End If
       SQ_OPER = 1
       PUB_TIPREG = 333
       pu_zona = nuc_repo!CLI_SUBGRUPO
       PUB_NUMTAB = pu_zona
       PUB_CODCIA = LK_CODCIA
       LEER_TAB_LLAVE
       If F1 > 10 Then F1 = F1 + 2
       xl.Cells(F1, 1) = ""
       If tab_llave.EOF Then
          xl.Cells(F1, 2) = "......."
       Else
          xl.Cells(F1, 2) = Trim(tab_llave!tab_nomlargo)
       End If
       wranF = "B" & F1 & ":B" & F1
       xl.Range(wranF).Font.Bold = True
       xl.Range(wranF).Font.Size = 12
       F1 = F1 + 1
       FILA_INICIAL = F1
    End If

  WS_SUBGRUPO = nuc_repo!CLI_SUBGRUPO


    N = 0
    total_cob = 0
    total_ven = 0
    TOTAL_SALDO = 0
    TOTAL_VEN_DIA = 0
    FRMIMP.lblproceso.Caption = Trim(nuc_repo!cli_nombre)
    DoEvents
    FRMIMP.ProgBar.Value = FRMIMP.ProgBar.Value + 1
    PSCAA(0) = LK_CODCIA
    PSCAA(1) = "C"
    PSCAA(2) = nuc_repo!cli_codclie
    PSCAA(3) = LK_FECHA_DIA
    CAA_REPO.Requery
    pu_codcia = LK_CODCIA
    TOTAL_SALDO = CAR_TOT_CPX("C", pu_codcia, nuc_repo!cli_codclie)

    
    
N = 0
Do Until N = 10
   N = N + 1
   tab_importe(N) = 0
   TAB_CONCEPTO(N) = ""
   tab_kilos(N) = 0
   tab_precio(N) = 0
Loop

    PSFAR(0) = LK_CODCIA
    PSFAR(1) = "C"
    PSFAR(2) = nuc_repo!cli_codclie
    PSFAR(3) = LOC_FECHA_INC
    far_r.Requery
    If far_r.EOF Then GoTo pasa
    WS_PRECIO = far_r!far_PRECIO
    N = 0
    wt_kilos = 0
    Do Until far_r.EOF
       If WS_PRECIO <> far_r!far_PRECIO Then
          N = N + 1
          tab_kilos(N) = wt_kilos
          tab_precio(N) = WS_PRECIO
          wt_kilos = 0
       End If
       wt_kilos = wt_kilos + far_r!far_PESO
       total_ven = redondea(far_r!far_PESO * far_r!far_PRECIO) + total_ven
       WS_PRECIO = far_r!far_PRECIO
       far_r.MoveNext
    Loop
       N = N + 1
       tab_kilos(N) = wt_kilos
       tab_precio(N) = WS_PRECIO
    
pasa:
    
N = 0
Do Until CAA_REPO.EOF
   If (Left(CAA_REPO!caa_concepto, 3)) = "Ext" Then
      N = N + 1
      tab_importe(N) = CAA_REPO!CAA_IMPORTE
   End If
   CAA_REPO.MoveNext
Loop

N = 0
CAA_REPO.MoveFirst
Do Until CAA_REPO.EOF
If CAA_REPO!CAA_IMPORTE = 555.1 Or CAA_REPO!CAA_IMPORTE = -555.1 Then
 Print " "
 End If

            If (Left(CAA_REPO!caa_concepto, 3)) = "Ext" Or (Left(CAA_REPO!caa_concepto, 3)) = "N/A" Or (CAA_REPO!CAA_SIGNO_CAJA > 1 And CAA_REPO!CAA_IMPORTE < 0) Or ((Left(CAA_REPO!caa_concepto, 8)) = "A Cuenta" And CAA_REPO!CAA_IMPORTE > 0) Then
               GoTo OTRO
            End If

         i = 1
         Do Until tab_importe(i) = 0
            If CAA_REPO!CAA_IMPORTE * -1 = tab_importe(i) Then
               tab_importe(i) = 0.001
               GoTo OTRO
            End If
            i = i + 1
         Loop

      
      
      If CAA_REPO!CAA_IMPORTE < 0 Then
         total_cob = CAA_REPO!CAA_IMPORTE + total_cob
      End If
      
      If (Nulo_Valor0(CAA_REPO!CAA_TIPMOV) <> 0 And LK_FECHA_DIA = CAA_REPO!CAA_FECHA And CAA_REPO!CAA_TIPDOC = "FA") Then
         TOTAL_VEN_DIA = CAA_REPO!CAA_IMPORTE + TOTAL_VEN_DIA
      End If
      
   
OTRO:
'   TOTAL_SALDO = caa_repo!CAA_SALDO
   
   CAA_REPO.MoveNext
Loop
Siguef:
   xcuenta = xcuenta + 1
   xl.Cells(F1, 1) = xcuenta
   xl.Cells(F1, 2) = Trim(nuc_repo!cli_nombre)
   If total_ven <> 0 Then xl.Cells(F1, 3) = tab_kilos(1)
   If total_ven <> 0 Then xl.Cells(F1, 4) = tab_precio(1)
   If total_ven <> 0 Then xl.Cells(F1, 5) = total_ven
   
   If total_cob <> 0 Then xl.Cells(F1, 6) = total_cob
   If (total_ven <> 0 Or total_cob <> 0) Then xl.Cells(F1, 7) = redondea(total_ven + total_cob)
   If (TOTAL_SALDO <> 0 Or TOTAL_VEN_DIA <> 0) Then
      If TOTAL_SALDO > TOTAL_VEN_DIA Then
         xl.Cells(F1, 8) = TOTAL_SALDO - TOTAL_VEN_DIA
      Else
         xl.Cells(F1, 9) = TOTAL_SALDO - TOTAL_VEN_DIA
      End If
   End If
   
   If TOTAL_VEN_DIA <> 0 Then xl.Cells(F1, 10) = TOTAL_VEN_DIA
   F1 = F1 + 1
   i = 2
   Do Until tab_kilos(i) = 0
      xl.Cells(F1, 3) = tab_kilos(i)
      xl.Cells(F1, 4) = tab_precio(i)
      F1 = F1 + 1
      i = i + 1
   Loop
   
   nuc_repo.MoveNext
Loop


Screen.MousePointer = 0
GoSub TOTALES
F1 = F1 + 2
total_ven = 0
Do Until total_ven = 1
   Mensaje = InputBox("ingrese nombre:")
   titulo = InputBox("ingrese importe:")
   If Len(Mensaje) = 0 Then
      total_ven = 1
   Else
      F1 = F1 + 1
      xl.Cells(F1, 2) = Mensaje
      xl.Cells(F1, 6) = titulo
   End If
Loop
      
   
GoTo VAMOS

WEXCEL:
  Dim xlchart As Chart
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FRMIMP.lblproceso.Caption = "Abriendo , Archivo Cobranzas.xls . . . "
  DoEvents
  WPAS = "131296"
  xl.Workbooks.Open "C:\ADMIN\OFFICE\COBRANZAS.xls", 0, True, 4, WPAS, WPAS
Return

VAMOS:
  DoEvents
  FRMIMP.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  DoEvents
  
  
  xl.Visible = True
  wranF = "F2:F4"
  xl.Range(wranF).Borders.Item(xlEdgeLeft).LineStyle = 2
  wranF = "J2:J4"
  xl.Range(wranF).Borders.Item(xlEdgeLeft).LineStyle = 2
  wranF = "A5:J" & F1
  xl.Range(wranF).Borders.LineStyle = 3
  xl.Cells(1, 2) = "R: " & PUB_CODTRA
  xl.Worksheets("Hoja1").Range("A1:X51").Locked = True
  xl.Worksheets("Hoja1").Protect "131296"
  xl.Worksheets(1).Activate
  xl.Application.Visible = True
  DoEvents
  FRMIMP.lblproceso.Visible = False
  FRMIMP.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FRMIMP.Pantalla.Enabled = True
  FRMIMP.Pantalla.Caption = "Por &Pantalla"
  FRMIMP.lblproceso.Visible = False

Exit Sub

GENERAL:
  F1 = F1 + 1
  xl.Cells(2, 3) = " Venta del dia :  " & DatePart("d", LOC_FECHA_INC) & "/" & DatePart("m", LOC_FECHA_INC) & "/90"
  xl.Cells(2, 8) = " Cobranzas del dia  :  " & DatePart("d", LK_FECHA_DIA) & "/" & DatePart("m", LK_FECHA_DIA) & "/90"
  xl.Cells(4, 10) = LK_FECHA_DIA
  xl.Cells(F1, 2) = "TOTAL GENERAL"
  xl.Cells(F1, 3) = COL3
  xl.Cells(F1, 5) = COL5
  xl.Cells(F1, 6) = COL6
  xl.Cells(F1, 7) = COL7
  xl.Cells(F1, 8) = COL8
  xl.Cells(F1, 9) = COL9
  xl.Cells(F1, 10) = COL10
  
  wranF = "B" & F1 & ":B" & F1
  xl.Range(wranF).Font.Bold = True
Return


'*** RUTINAS PARA IMPRIMIR



WPROGRESO:

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
Return


Return


Exit Sub
CANCELA:
  FRMIMP.Pantalla.Enabled = True
  FRMIMP.Pantalla.Caption = "Por &Pantalla"
  FRMIMP.lblproceso.Visible = False
  LOC_CANCELA = 0
  Set xl = Nothing
  Screen.MousePointer = 0

Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FRMIMP

TOTALES:
  xl.Cells(F1, 2) = "              SubTotal:"
  wran1 = LETRAS(10) & FILA_INICIAL
  wran2 = LETRAS(10) & F1 - 1
  wranF = LETRAS(10) & F1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  COL10 = COL10 + xl.Range(wranF)
  
  wran1 = LETRAS(9) & FILA_INICIAL
  wran2 = LETRAS(9) & F1 - 1
  wranF = LETRAS(9) & F1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  COL9 = COL9 + xl.Range(wranF)
  
  wran1 = LETRAS(8) & FILA_INICIAL
  wran2 = LETRAS(8) & F1 - 1
  wranF = LETRAS(8) & F1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  COL8 = COL8 + xl.Range(wranF)
  
  wran1 = LETRAS(7) & FILA_INICIAL
  wran2 = LETRAS(7) & F1 - 1
  wranF = LETRAS(7) & F1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  COL7 = COL7 + xl.Range(wranF)

  wran1 = LETRAS(6) & FILA_INICIAL
  wran2 = LETRAS(6) & F1 - 1
  wranF = LETRAS(6) & F1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  COL6 = COL6 + xl.Range(wranF)
  
  wran1 = LETRAS(5) & FILA_INICIAL
  wran2 = LETRAS(5) & F1 - 1
  wranF = LETRAS(5) & F1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  COL5 = COL5 + xl.Range(wranF)
  
  wran1 = LETRAS(3) & FILA_INICIAL
  wran2 = LETRAS(3) & F1 - 1
  wranF = LETRAS(3) & F1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  COL3 = COL3 + xl.Range(wranF)
  F1 = F1 + 1

  Return

End Sub
Public Sub PRO_REPORTE(WDIS As String)
Dim wformula, wformula1, wformula2, wformula3, wformula4
Dim DIA, MES, ANO
Dim DIA1, MES1, ANO1
Dim PU_MONEDA As String * 1
'On Error GoTo SALE
Pantalla.Enabled = False
cerrar.Enabled = False
If Trim(UCase(Wfile)) = "LOTE_IMP" Then
  LOC_NUCLEO = Left(rLista.Text, 1)
  If WDIS = "P" Then
   LKCHEK = True
  ElseIf WDIS = "I" Then
    LKCHEK = True
  End If
  'LOTE_IMP  LOC_NUCLEO, 99
  Pantalla.Enabled = True
  cerrar.Enabled = True
  Exit Sub
End If

If Trim(UCase(Wfile)) = "IMP_FACTURACION" Then
 If Trim(CMBFGB.Text) = "" Then
   MsgBox " Seleccionar Documento . .", 48, Pub_Titulo
   GoTo VAMOS
 End If
 If Trim(txtSerie.Text) = "" Then
   MsgBox " Seleccionar Serie .", 48, Pub_Titulo
   Azul txtSerie, txtSerie
   GoTo VAMOS
 End If
 If Trim(txtnumfac1.Text) = "" Then
   MsgBox " Seleccionar Numero Inicial .", 48, Pub_Titulo
   Azul txtnumfac1, txtnumfac1
   GoTo VAMOS
 End If
 If Trim(txtnumfac2.Text) = "" Then
   MsgBox " Seleccionar Numero final .", 48, Pub_Titulo
   Azul txtnumfac2, txtnumfac2
   GoTo VAMOS
 End If
End If

  If Trim(UCase(Wfile)) = "CLI_IMP" Or Trim(UCase(Wfile)) = "VENTAS_IMP" Or Trim(UCase(Wfile)) = "ABONOS_IMP" Or Trim(UCase(Wfile)) = "CLI_SALDOS" Or Trim(UCase(Wfile)) = "VENTASMES_IMP" Or Trim(UCase(Wfile)) = "VENTAS_POLLOS" Or Trim(UCase(Wfile)) = "VENTAS_DIA_IMP" Or Trim(UCase(Wfile)) = "FORMULA_IMP" Or Trim(UCase(Wfile)) = "CLI_COBRANZAS" Or Trim(UCase(Wfile)) = "PORPAGAR" Or Trim(UCase(Wfile)) = "CONVERSION" Or Trim(UCase(Wfile)) = "RECORD_VENDEDOR" Or Trim(UCase(Wfile)) = "MORTAL" Or Trim(UCase(Wfile)) = "IMP_FACTURACION" Or Trim(UCase(Wfile)) = "PORPAGAR2" Or Trim(UCase(Wfile)) = "COMISIONES" Or Trim(UCase(Wfile)) = "IMP_LOTES" Or Trim(UCase(Wfile)) = "CLI_SABADOS" Or Trim(UCase(Wfile)) = "CLI_SERVICIOS" Or Trim(UCase(Wfile)) = "REPO_CAJA_GEN" Or Trim(UCase(Wfile)) = "CLI_SERV_RESU" Or Trim(UCase(Wfile)) = "PORCOBRAR2" Or Trim(UCase(Wfile)) = "REPO_CAJA_CHICA" Or Trim(UCase(Wfile)) = "CAJA_SERVICIOS" Or Trim(UCase(Wfile)) = "CTA_SERVICIOS" Then
  LOC_FECHA_INC = Trim(txtfecha.Text)
  
  If Trim(UCase(Wfile)) = "REPO_CAJA_GEN" Then
     REPO_CAJA_GEN
  End If
    If Trim(UCase(Wfile)) = "REPO_CAJA_CHICA" Then
     REPO_CAJA_CHICA
  End If


  Pantalla.Enabled = True
  FRMIMP.ProgBar.Visible = False
  cerrar.Enabled = True
  Exit Sub
End If


Screen.MousePointer = 11
FRMIMP.ProgBar.Min = 0
FRMIMP.ProgBar.Max = 10
FRMIMP.ProgBar.Value = 0
FRMIMP.ProgBar.Visible = True
DoEvents
FRMIMP.lblproceso.Visible = True
DoEvents

If Len(Wfile) <> 0 Then
  FRMIMP.ProgBar.Value = 2
  Reportes.ReportFileName = PUB_RUTA_REPORTE & Wfile
  Reportes.WindowTitle = "Reporte :  " & Trim(tra_llave(1))
  FRMIMP.ProgBar.Value = 4
  DoEvents
  If WDIS = "P" Then
   Reportes.Destination = crptToWindow
  ElseIf WDIS = "I" Then
   Reportes.Destination = crptToPrinter
  End If
  Reportes.WindowLeft = 2
  Reportes.WindowTop = 70
  Reportes.WindowWidth = 635
  Reportes.WindowHeight = 390
  Reportes.Formulas(0) = ""
  Reportes.Formulas(1) = ""
  Reportes.Formulas(2) = ""
  Reportes.Formulas(3) = ""
  Reportes.Formulas(4) = ""
  FRMIMP.ProgBar.Value = 6
  WS1 = "01/" & Format(LK_FECHA_DIA, "mm") & "/" & Format(LK_FECHA_DIA, "yyyy")
  WMES = CDate(WS1)
  DIA1 = Day(WMES)
  MES1 = Month(WMES)
  ANO1 = Year(WMES)

  
  DIA = Day(LK_FECHA_DIA)
  MES = Month(LK_FECHA_DIA)
  ANO = Year(LK_FECHA_DIA)
  If UCase(Trim(Wfile)) = "LOTES.RPT" Then
     wformula = "FECHA=  '" & Format(LK_FECHA_DIA, "dddd, d mmmm yyyy") & "'"
     wformula1 = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
     Reportes.Formulas(0) = wformula
     Reportes.Formulas(1) = wformula1
     Reportes.SelectionFormula = "{FACART.FAR_ESTADO} <> 'E' and {FACART.FAR_FECHA} = Date ( " & ANO & "," & MES & "," & DIA & ") and {FACART.FAR_CODCIA} = '" & LK_CODCIA & "'"
     FRMIMP.ProgBar.Value = 9
     GoTo MUESTRA_REPO
   End If
   If UCase(Trim(Wfile)) = "INSUMOS.RPT" Then
     wformula = "FECHA=  '" & Format(LK_FECHA_DIA, "dddd, d mmmm yyyy") & "'"
     wformula1 = "F_ingreso= if {FACART.FAR_SIGNO_ARM} = 1  and  {FACART.FAR_FECHA} = Date (" & ANO & "," & MES & "," & DIA & ") then {FACART.FAR_CANTIDAD}"
     wformula2 = "F_salida= if {FACART.FAR_SIGNO_ARM} = -1  and  {FACART.FAR_FECHA} = Date (" & ANO & "," & MES & "," & DIA & ") then {FACART.FAR_CANTIDAD}"
     wformula3 = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
     wformula4 = "CONSIS= " & PROTER_SALDO
     Reportes.SelectionFormula = "{FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {TABLAS.TAB_TIPREG} = 123 and {TABLAS.TAB_CODCIA} = '" & LK_CODCIA & "' And {FACART.FAR_ESTADO} <> 'E' And {ARTI.ART_FAMILIA} = 12"
     Reportes.Formulas(0) = wformula
     Reportes.Formulas(1) = wformula1
     Reportes.Formulas(2) = wformula2
     Reportes.Formulas(3) = wformula3
     Reportes.Formulas(4) = wformula4
     GoTo MUESTRA_REPO
   End If
     If UCase(Trim(Wfile)) = "SALDODIA.RPT" Then
     wformula = "FECHA=  '" & Format(LK_FECHA_DIA, "dddd, d mmmm yyyy") & "'"
     wformula1 = "F_ingreso= if {FACART.FAR_SIGNO_ARM} = 1  and  {FACART.FAR_FECHA} = Date (" & ANO & "," & MES & "," & DIA & ")  then {FACART.FAR_CANTIDAD}"
     wformula2 = "F_salida= if {FACART.FAR_SIGNO_ARM} = -1  and  {FACART.FAR_FECHA} = Date (" & ANO & "," & MES & "," & DIA & ")  then {FACART.FAR_CANTIDAD}"
     wformula3 = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
     PU_MONEDA = UCase(Trim(txtNucleo.Text))
     Reportes.SelectionFormula = "{FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {TABLAS.TAB_TIPREG} = 122 and {TABLAS.TAB_CODCIA} = '" & LK_CODCIA & "'AND {ARTI.ART_FAMILIA} = 1  AND {ARTI.ART_MONEDA} = '" & PU_MONEDA & "' "
     Reportes.Formulas(0) = wformula
     Reportes.Formulas(1) = wformula1
     Reportes.Formulas(2) = wformula2
     Reportes.Formulas(3) = wformula3
     GoTo MUESTRA_REPO
   End If
      If UCase(Trim(Wfile)) = "CONSUDOL.RPT" Then
     wformula = "FECHA=  '" & Format(LK_FECHA_DIA, "dddd, d mmmm yyyy") & "'"
     wformula1 = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
     wformula2 = "GUIA=  '" & txtNucleo.Text & "'"
     Reportes.SelectionFormula = "{FACART.FAR_TIPMOV} = 90 AND {FACART.FAR_FECHA} = Date ( " & ANO & "," & MES & "," & DIA & ") AND {ARTI.ART_MONEDA} = 'D' and {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_ESTADO} <> 'E' "
     Reportes.Formulas(0) = wformula
     Reportes.Formulas(1) = wformula1
     Reportes.Formulas(2) = wformula2
     GoTo MUESTRA_REPO
   End If
   If UCase(Trim(Wfile)) = "CONSUSOL.RPT" Then
     wformula = "FECHA=  '" & Format(LK_FECHA_DIA, "dddd, d mmmm yyyy") & "'"
     wformula1 = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
     wformula2 = "GUIA=  '" & txtNucleo.Text & "'"
     Reportes.SelectionFormula = "{FACART.FAR_TIPMOV} = 90 AND {FACART.FAR_FECHA} = Date ( " & ANO & "," & MES & "," & DIA & ") AND {ARTI.ART_MONEDA} = 'S' and {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_ESTADO} <> 'E' "
     Reportes.Formulas(0) = wformula
     Reportes.Formulas(1) = wformula1
     Reportes.Formulas(2) = wformula2
     GoTo MUESTRA_REPO
   End If
     If UCase(Trim(Wfile)) = "ACU.RPT" Then
     wformula = "FECHA=  '" & Format(LK_FECHA_DIA, "dddd, d mmmm yyyy") & "'"
     wformula1 = "F_ingreso= if {FACART.FAR_SIGNO_ARM} = 1  and  {FACART.FAR_FECHA} >= Date (" & ANO1 & "," & MES1 & "," & DIA1 & ") and  {FACART.FAR_FECHA} <= Date (" & ANO & "," & MES & "," & DIA & ") then {FACART.FAR_CANTIDAD}"
     wformula2 = "F_salida= if {FACART.FAR_SIGNO_ARM} = -1  and  {FACART.FAR_FECHA} >= Date (" & ANO1 & "," & MES1 & "," & DIA1 & ") and  {FACART.FAR_FECHA} <= Date (" & ANO & "," & MES & "," & DIA & ") then {FACART.FAR_CANTIDAD}"
     wformula3 = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
     PU_MONEDA = UCase(Trim(txtNucleo.Text))
     Reportes.SelectionFormula = "{FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {TABLAS.TAB_TIPREG} = 122 and {TABLAS.TAB_CODCIA} = '" & LK_CODCIA & "' AND {ARTI.ART_MONEDA} = '" & PU_MONEDA & "' "
     Reportes.Formulas(0) = wformula
     Reportes.Formulas(1) = wformula1
     Reportes.Formulas(2) = wformula2
     Reportes.Formulas(3) = wformula3
     GoTo MUESTRA_REPO
    End If
   If UCase(Trim(Wfile)) = "MOVIDIA.RPT" Then
     wformula = "FECHA=  '" & Format(LK_FECHA_DIA, "dddd, d mmmm yyyy") & "'"
     Reportes.SelectionFormula = "{FACART.FAR_FECHA} = Date ( " & ANO & "," & MES & "," & DIA & ") and {TABLAS.TAB_CODCIA} = '00' and {TABLAS.TAB_TIPREG} = 4 and {FACART.FAR_TIPMOV} = 20 and {ARTI.ART_FAMILIA} = 12  and {FACART.FAR_ESTADO} <> 'E' and {FACART.FAR_CODCIA} = '" & LK_CODCIA & "'"
     Reportes.Formulas(0) = wformula
     GoTo MUESTRA_REPO
   End If
   If UCase(Trim(Wfile)) = "INVENT.RPT" Then
     wformula = "FECHA=  '" & Format(LK_FECHA_DIA, "dddd, d mmmm yyyy") & "'"
     Reportes.SelectionFormula = "{FACART.FAR_FECHA} = Date ( " & ANO & "," & MES & "," & DIA & ") and {TABLAS.TAB_CODCIA} = '00' and {TABLAS.TAB_TIPREG} = 4 and {FACART.FAR_TIPMOV} <> 20 and {ARTI.ART_FAMILIA} = 12 and {FACART.FAR_TIPMOV} <> 90 and {FACART.FAR_TIPMOV} <> 0 and {FACART.FAR_CODCIA} = '" & LK_CODCIA & "'"
     Reportes.Formulas(0) = wformula
     GoTo MUESTRA_REPO
   End If
   
   If UCase(Trim(Wfile)) = "BALANCE.RPT" Then
     wformula = "FECHA=  '" & Format(LK_FECHA_DIA, "dddd, d mmmm yyyy") & "'"
     Reportes.SelectionFormula = "{COMAEST.COM_ACT_PAS} <> 0 and {TABLAS.TAB_TIPREG} = 16 and {COMAEST.COM_CODCIA} = '" & LK_CODCIA & "'"
     Reportes.Formulas(0) = wformula
     GoTo MUESTRA_REPO
   End If
   If UCase(Trim(Wfile)) = "DIARIO.RPT" Then
     wformula = "FECHA=  '" & Format(LK_FECHA_DIA, "dddd, d mmmm yyyy") & "'"
     Reportes.Formulas(0) = wformula
     GoTo MUESTRA_REPO
   End If
   If UCase(Trim(Wfile)) = "ESTADOGP" Or UCase(Trim(Wfile)) = "ESTADOGP_2" Then
        If WDIS = "P" Then
         LKCHEK = True
         ElseIf WDIS = "I" Then
            LKCHEK = False
        End If
        If UCase(Trim(Wfile)) = "ESTADOGP" Then
            POWER_REPORT 77
        Else
            POWER_REPORT 78
        End If
          
        Pantalla.Enabled = True
        cerrar.Enabled = True
        Screen.MousePointer = 0
        FRMIMP.ProgBar.Visible = False
        Exit Sub
       wformula = "FECHA=  '" & Format(LK_FECHA_DIA, "dddd, d mmmm yyyy") & "'"
       Reportes.SelectionFormula = "{TABLAS.TAB_TIPREG} = 16 and {COMAEST.COM_TIPO_CTA} in 6 to 7 and {COMAEST.COM_CODCIA} = '" & LK_CODCIA & "'"
       Reportes.Formulas(0) = wformula
     GoTo MUESTRA_REPO
   End If
   If UCase(Trim(Wfile)) = "CUENTAS.RPT" Or UCase(Trim(Wfile)) = "CTAMAEST.RPT" Then
     wformula = "FECHA=  '" & Format(LK_FECHA_DIA, "dddd, d mmmm yyyy") & "'"
     wformula1 = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
     Reportes.SelectionFormula = "{COMAEST.COM_CODCIA} = '" & LK_CODCIA & "'"
     Reportes.Formulas(0) = wformula
     Reportes.Formulas(1) = wformula1
     GoTo MUESTRA_REPO
   End If
   If UCase(Trim(Wfile)) = "DESPACHO.RPT" Then
     wformula = "FECHA=  '" & Format(LK_FECHA_DIA, "dddd, d mmmm yyyy") & "'"
     wformula1 = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
     wformula2 = "FECHA_SISTEM= {FACART.FAR_FECHA} = Date ( " & ANO & "," & MES & "," & DIA & ")"
     Reportes.SelectionFormula = "{CLIENTES.CLI_CUENTA_CONTAB} <> '             ' and  {CLIENTES.CLI_estado} ='A' and {FACART.FAR_TIPMOV} = 80 and {FACART.FAR_ESTADO}<> 'E' and {FACART.FAR_CODCIA} = '" & LK_CODCIA & "'"
     Reportes.Formulas(0) = wformula
     Reportes.Formulas(1) = wformula1
     Reportes.Formulas(2) = wformula2
     GoTo MUESTRA_REPO
   End If
   If UCase(Trim(Wfile)) = "FACVEN.RPT" Then
     wformula = "FECHA=  '" & Format(LK_FECHA_DIA, "dddd, d mmmm yyyy") & "'"
     wformula1 = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
     Reportes.SelectionFormula = "{CARTERA.CAR_IMPORTE} <> 0 and {CLIENTES.CLI_CP} = 'P' and {CLIENTES.CLI_CODCIA} = '" & LK_CODCIA & "' and {CARTERA.CAR_IMPORTE} >= 1.00 and {CARTERA.CAR_FECHA_VCTO} <= Date ( " & ANO & "," & MES & "," & DIA & ")"
     Reportes.Formulas(0) = wformula
     Reportes.Formulas(1) = wformula1
     GoTo MUESTRA_REPO
   End If
   If UCase(Trim(Wfile)) = "VENDIA.RPT" Then
     LOC_FECHA_INC = Trim(txtfecha.Text)
      If Not IsDate(LOC_FECHA_INC) Then
         MsgBox "Fecha Invalidad ...", 48, Pub_Titulo
         GoTo VAMOS
     End If
     DIA1 = Day(LOC_FECHA_INC)
     MES1 = Month(LOC_FECHA_INC)
     ANO1 = Year(LOC_FECHA_INC)
     wformula = "FECHA=  '" & Format(LK_FECHA_DIA, "dddd, d mmmm yyyy") & "'"
     wformula1 = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
     Reportes.SelectionFormula = "{FACART.FAR_FECHA} >= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ") AND {FACART.FAR_TIPMOV} = 28 and {FACART.FAR_CODCIA} = '" & LK_CODCIA & "'"
     Reportes.Formulas(0) = wformula
     Reportes.Formulas(1) = wformula1
     GoTo MUESTRA_REPO
   End If
End If

MUESTRA_REPO:
   Reportes.Action = 1
   FRMIMP.ProgBar.Value = 10
VAMOS:
Screen.MousePointer = 0
FRMIMP.ProgBar.Visible = False


FRMIMP.lblproceso.Visible = False
Pantalla.Enabled = True
cerrar.Enabled = True

Exit Sub

SALE:
 Screen.MousePointer = 0
 MsgBox Err.Description, 48, Pub_Titulo
 Pantalla.Enabled = True
 cerrar.Enabled = True

End Sub

Public Sub CLI_SABADOS()
On ERRO GoTo FINTODO
Dim success%
Dim wranF, wran1, wran2, WPAS
Dim LETRAS(24) As String * 1
Dim MESES(12) As String
Dim WS_F1, WS_F2, WS_F3, WS_F4 As Date
Dim WW_F1, WW_F2, WW_F3, WW_F4 As Date

Dim FECHA_INICIO
Dim WSFECHA As Date
Dim xMES1, xMES2 As Integer
Dim con  As Integer
Dim TABLAGEN() As String * 30
Dim COL9, COL8, COL7, COL6, COL5, COL4, COL3, COL10
Dim TABLAGEN2() As String * 30
Dim i, WS_SUBGRUPO
Dim wt_kilos As Currency
Dim WS_PRECIO As Currency
Dim todos_cli As Integer
Dim filas_cli As Integer
Dim xpollos As String * 6
Dim xdias As String * 6
Dim C1 As Integer
Dim F1 As Integer
Dim WTABULA2 As String * 1
Dim WTABULA As String * 1
Dim RUTA As String
Dim WR_FECHA As String
Dim WR_LOTE As String
Dim WR_DAT1 As String
Dim SUM_cantidad As Currency
Dim WR_soles As String * 10
Dim WR_CIA As String
Dim cade As String
Dim tab_importe(10) As Currency
Dim TAB_CONCEPTO(10) As String * 30
Dim tab_kilos(10) As Currency
Dim tab_precio(10) As Currency

Dim SUM_SALDOS As Currency
Dim w_cant As String * 12
Dim w_sol As String * 10
Dim WCANT As Currency

Dim nuc_repo  As rdoResultset
Dim PSNUC_REPO As rdoQuery
Dim far_r  As rdoResultset
Dim PSFAR As rdoQuery

Dim Mensaje, titulo, valorpred As String
Dim wvalor
Dim valor As Double
Dim xcuenta As Integer
Dim PAG_HOJAS As Integer
Dim WS_SALDO As Currency
Dim WS_S As String
Dim ws_tt As String
Dim TOTAL_SALDO As Currency
'FRMIMP.pantalla.Caption = "Cancelar"
'FRMIMP.pantalla.Enabled = True
DoEvents
cade = "SELECT * FROM CARACU WHERE CAA_CODCIA = ? AND CAA_CP = ? AND CAA_CODCLIE = ? AND CAA_FECHA <= ?  ORDER BY CAA_CODCLIE, CAA_FECHA, CAA_NUM_OPER"
Set PSCAA = CN.CreateQuery("", cade)
Set CAA_REPO = PSCAA.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

cade = "SELECT * FROM CLIENTES WHERE CLI_CODCIA = ? AND CLI_CP = ? AND CLI_ESTADO = 'B'  ORDER BY CLI_SUBGRUPO, CLI_NOMBRE"
Set PSNUC_REPO = CN.CreateQuery("", cade)
Set nuc_repo = PSNUC_REPO.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

cade = "SELECT * FROM facart WHERE FAR_CODCIA = ? AND FAR_CP = ? AND FAR_CODCLIE = ? AND FAR_FECHA = ? AND FAR_TIPMOV = 28  AND FAR_ESTADO <> 'E' ORDER BY FAR_PRECIO"
Set PSFAR = CN.CreateQuery("", cade)
Set far_r = PSFAR.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

If wver = 99 Then FRMIMP.ProgBar.Value = 1
LOC_FECHA_INC = DateAdd("d", -1, LK_FECHA_DIA)

'*** LLENA DE TABLAS ... CLIENTES NUCLEOS
FECHA_INICIO = LOC_FECHA_INC
PSNUC_REPO(0) = LK_CODCIA
PSNUC_REPO(1) = "C"
Screen.MousePointer = 11
SUM_SALDOS = 0
nuc_repo.Requery
If nuc_repo.EOF Then
  Screen.MousePointer = 0
  MsgBox "NO EXISTE NINGUNO ..", 48, Pub_Titulo
  FRMIMP.ProgBar.Visible = False
  FRMIMP.lblproceso.Visible = False
  Exit Sub
End If
FRMIMP.ProgBar.Min = 0
FRMIMP.ProgBar.Max = nuc_repo.RowCount
FRMIMP.ProgBar.Value = 0
FRMIMP.ProgBar.Visible = True
DoEvents
FRMIMP.lblproceso.Visible = True
FRMIMP.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
todos_cli = nuc_repo.RowCount
Dim wFECHASIGUE As Date
Dim wtC, wtF

GoSub WEXCEL
PAG_HOJAS = 1
F1 = 0
ws_fecha = LK_FECHA_DIA
Do Until F1 = 1
   If DatePart("w", ws_fecha) = 3 Then
      F1 = 1
   Else
      ws_fecha = DateAdd("d", -1, ws_fecha)
   End If
Loop
WS_F1 = ws_fecha
WS_F2 = DateAdd("d", -7, WS_F1)
WS_F3 = DateAdd("d", -7, WS_F2)
WS_F4 = DateAdd("d", -7, WS_F3)
   
F1 = 0
WW_F1 = LK_FECHA_DIA

WW_F2 = "01" & "/" & DatePart("m", LK_FECHA_DIA) & "/" & DatePart("yYYY", LK_FECHA_DIA)
WW_F2 = DateAdd("d", -1, WW_F2)
WW_F3 = "01" & "/" & DatePart("m", WW_F2) & "/" & DatePart("yYYY", WW_F2)
WW_F3 = DateAdd("d", -1, WW_F3)
WW_F4 = "01" & "/" & DatePart("m", WW_F3) & "/" & DatePart("YYYy", WW_F3)
WW_F4 = DateAdd("d", -1, WW_F4)

   
   
   
xl.Worksheets(1).Activate
'GoSub LETRAS
WS_SALDO = 0
xcuenta = 0
F1 = 5
WS_SUBGRUPO = ""
FILA_INICIAL = 5
Do Until nuc_repo.EOF ' loop 1
    If nuc_repo!CLI_SUBGRUPO <> WS_SUBGRUPO Then
'       If WS_SUBGRUPO <> "" Then GoSub TOTALES
'       If nuc_repo!CLI_SUBGRUPO = 50 Then
'          GoSub GENERAL
'       End If
       SQ_OPER = 1
       PUB_TIPREG = 333
       pu_zona = nuc_repo!CLI_SUBGRUPO
       PUB_NUMTAB = pu_zona
       PUB_CODCIA = LK_CODCIA
       LEER_TAB_LLAVE
       F1 = F1 + 1
       xl.Cells(F1, 1) = ""
       If tab_llave.EOF Then
          xl.Cells(F1, 2) = "......."
       Else
          xl.Cells(F1, 2) = Trim(tab_llave!tab_nomlargo)
       End If
       wranF = "B" & F1 & ":B" & F1
       xl.Range(wranF).Font.Bold = True
       xl.Range(wranF).Font.Size = 12
    End If

  WS_SUBGRUPO = nuc_repo!CLI_SUBGRUPO


    N = 0
    total_cob = 0
    total_ven = 0
    TOTAL_SALDO = 0
    TOTAL_VEN_DIA = 0
    FRMIMP.lblproceso.Caption = Trim(nuc_repo!cli_nombre)
    DoEvents
    pu_codcia = LK_CODCIA
   If nuc_repo!cli_codclie = 17735 Then
       Print ""
    End If

    TOTAL_SALDO = CAR_TOT_CPX("C", pu_codcia, nuc_repo!cli_codclie)
    If TOTAL_SALDO <= 0 Then GoTo otro_cli
    
    FRMIMP.ProgBar.Value = FRMIMP.ProgBar.Value + 1
    If nuc_repo!cli_codclie = 17735 Then
       Print ""
    End If
    
WS_SALDO1 = 0
WS_SALDO2 = 0
WS_SALDO3 = 0

WS_SALDO4 = 0
WW_SALDO1 = 0
WW_SALDO2 = 0
WW_SALDO3 = 0
WW_SALDO4 = 0
FX1 = 0
FX2 = 0
FX3 = 0
FX4 = 0
XX1 = 0
XX2 = 0
XX3 = 0
XX4 = 0

PSCAA(3) = WS_F1
GoSub BUSCA
WS_SALDO1 = WS_SALDO
If WS_SALDO <= 0 Then GoTo otro_cli

PSCAA(3) = WS_F2
GoSub BUSCA
WS_SALDO2 = WS_SALDO

PSCAA(3) = WS_F3
GoSub BUSCA
WS_SALDO3 = WS_SALDO

PSCAA(3) = WS_F4
GoSub BUSCA
WS_SALDO4 = WS_SALDO

PSCAA(3) = WW_F1
'GoSub BUSCA
WW_SALDO1 = WS_SALDO

PSCAA(3) = WW_F2
'GoSub BUSCA
WW_SALDO2 = WS_SALDO

PSCAA(3) = WW_F3
'GoSub BUSCA
WW_SALDO3 = WS_SALDO

PSCAA(3) = WW_F4
'GoSub BUSCA
WW_SALDO4 = WS_SALDO

Siguef:
   F1 = F1 + 1
   xcuenta = xcuenta + 1
   xl.Cells(F1, 1) = xcuenta
   xl.Cells(F1, 2) = Trim(nuc_repo!cli_nombre)
   
   xl.Cells(F1, 3) = WS_SALDO4
   xl.Cells(F1, 4) = WS_SALDO3
   xl.Cells(F1, 5) = WS_SALDO2
   xl.Cells(F1, 6) = WS_SALDO1
 
 '  xl.Cells(F1, 7) = WW_SALDO4
 '  xl.Cells(F1, 8) = WW_SALDO3
 '  xl.Cells(F1, 9) = WW_SALDO2
 '  xl.Cells(F1, 10) = WW_SALDO1
   
otro_cli:
   nuc_repo.MoveNext
Loop

Screen.MousePointer = 0
'GoSub TOTALES
F1 = F1 + 2
xl.Cells(4, 3) = WS_F4
xl.Cells(4, 4) = WS_F3
xl.Cells(4, 5) = WS_F2
xl.Cells(4, 6) = WS_F1
      
      
GoTo VAMOS

WEXCEL:
  Dim xlchart As Chart
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FRMIMP.lblproceso.Caption = "Abriendo , Archivo Cobranzas.xls . . . "
  DoEvents
  WPAS = "131296"
  xl.Workbooks.Open "C:\ADMIN\OFFICE\SABADOS.xls", 0, True, 4, WPAS, WPAS
Return

VAMOS:
  DoEvents
  FRMIMP.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  DoEvents
  
  
  xl.Visible = True
'  wranF = "F2:F4"
'  xl.Range(wranF).Borders.Item(xlEdgeLeft).LineStyle = 2
'  wranF = "J2:J4"
'  xl.Range(wranF).Borders.Item(xlEdgeLeft).LineStyle = 2
  wranF = "A5:F" & F1
  xl.Range(wranF).Borders.LineStyle = 2
  xl.Cells(1, 2) = "R: " & PUB_CODTRA
  xl.Worksheets("Hoja1").Range("A1:X51").Locked = True
  xl.Worksheets("Hoja1").Protect "131296"
  xl.Worksheets(1).Activate
  xl.Application.Visible = True
  DoEvents
  FRMIMP.lblproceso.Visible = False
  FRMIMP.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FRMIMP.Pantalla.Enabled = True
  FRMIMP.Pantalla.Caption = "Por &Pantalla"
  FRMIMP.lblproceso.Visible = False

Exit Sub

GENERAL:
  F1 = F1 + 1
  xl.Cells(2, 3) = " Venta del dia :  " & DatePart("d", LOC_FECHA_INC) & "/" & DatePart("m", LOC_FECHA_INC) & "/90"
  xl.Cells(2, 8) = " Cobranzas del dia  :  " & DatePart("d", LK_FECHA_DIA) & "/" & DatePart("m", LK_FECHA_DIA) & "/90"
  xl.Cells(4, 10) = LK_FECHA_DIA
  xl.Cells(F1, 2) = "TOTAL GENERAL"
  xl.Cells(F1, 3) = COL3
  xl.Cells(F1, 5) = COL5
  xl.Cells(F1, 6) = COL6
  xl.Cells(F1, 7) = COL7
  xl.Cells(F1, 8) = COL8
  xl.Cells(F1, 9) = COL9
  xl.Cells(F1, 10) = COL10
  
  wranF = "B" & F1 & ":B" & F1
  xl.Range(wranF).Font.Bold = True
Return


'*** RUTINAS PARA IMPRIMIR



WPROGRESO:

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
Return


Return


Exit Sub
CANCELA:
  FRMIMP.Pantalla.Enabled = True
  FRMIMP.Pantalla.Caption = "Por &Pantalla"
  FRMIMP.lblproceso.Visible = False
  LOC_CANCELA = 0
  Set xl = Nothing
  Screen.MousePointer = 0

Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FRMIMP

TOTALES:
  xl.Cells(F1, 2) = "              SubTotal:"
  wran1 = LETRAS(10) & FILA_INICIAL
  wran2 = LETRAS(10) & F1 - 1
  wranF = LETRAS(10) & F1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  COL10 = COL10 + xl.Range(wranF)
  
  wran1 = LETRAS(9) & FILA_INICIAL
  wran2 = LETRAS(9) & F1 - 1
  wranF = LETRAS(9) & F1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  COL9 = COL9 + xl.Range(wranF)
  
  wran1 = LETRAS(8) & FILA_INICIAL
  wran2 = LETRAS(8) & F1 - 1
  wranF = LETRAS(8) & F1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  COL8 = COL8 + xl.Range(wranF)
  
  wran1 = LETRAS(7) & FILA_INICIAL
  wran2 = LETRAS(7) & F1 - 1
  wranF = LETRAS(7) & F1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  COL7 = COL7 + xl.Range(wranF)

  wran1 = LETRAS(6) & FILA_INICIAL
  wran2 = LETRAS(6) & F1 - 1
  wranF = LETRAS(6) & F1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  COL6 = COL6 + xl.Range(wranF)
  
  wran1 = LETRAS(5) & FILA_INICIAL
  wran2 = LETRAS(5) & F1 - 1
  wranF = LETRAS(5) & F1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  COL5 = COL5 + xl.Range(wranF)
  
  wran1 = LETRAS(3) & FILA_INICIAL
  wran2 = LETRAS(3) & F1 - 1
  wranF = LETRAS(3) & F1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  COL3 = COL3 + xl.Range(wranF)
  F1 = F1 + 1

  Return
BUSCA:
    WS_SALDO = 0
    PSCAA(0) = LK_CODCIA
    PSCAA(1) = "C"
    PSCAA(2) = nuc_repo!cli_codclie
    CAA_REPO.Requery
    If CAA_REPO.EOF Then Return
    
CAA_REPO.MoveLast
N = 0
WS_SALDO = 0
Do Until CAA_REPO.BOF Or N = 1
   If CAA_REPO!CAA_IMPORTE > 0 And CAA_REPO!CAA_TIPDOC = "FA" And CAA_REPO!CAA_IMPORTE = CAA_REPO!CAA_TOTAL And CAA_REPO!CAA_FECHA = PSCAA(3) Then
   Else
      N = 1
      WS_SALDO = CAA_REPO!CAA_SALDO
   End If
   
   CAA_REPO.MovePrevious
Loop
Return

End Sub
Public Sub BALANCE()
On ERRO GoTo FINTODO
Dim success%
Dim wranF, wran1, wran2, WPAS


Dim bal_repo  As rdoResultset
Dim PSBAL_REPO As rdoQuery

Dim Mensaje, titulo, valorpred As String
Dim wvalor
Dim valor As Double
Dim xcuenta As Integer
Dim PAG_HOJAS As Integer
Dim WS_SALDO As Currency
Dim WS_S As String
Dim ws_tt As String
Dim TOTAL_SALDO As Currency

cade = "SELECT * FROM COMAEST WHERE CCOM_CODCIA = ?  ORDER BY COM_CUENTA, COM_TIPO_CTA"
Set PSBAL_REPO = CN.CreateQuery("", cade)
Set bal_repo = PSBAL.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

PSBAL_REPO(0) = LK_CODCIA
Screen.MousePointer = 11
bal_repo.Requery
If bal_repo.EOF Then
  Screen.MousePointer = 0
  MsgBox "NO EXISTE NINGUNO ..", 48, Pub_Titulo
  FRMIMP.ProgBar.Visible = False
  FRMIMP.lblproceso.Visible = False
  Exit Sub
End If
FRMIMP.ProgBar.Min = 0
FRMIMP.ProgBar.Max = bal_repo.RowCount
FRMIMP.ProgBar.Value = 0
FRMIMP.ProgBar.Visible = True
DoEvents
FRMIMP.lblproceso.Visible = True
FRMIMP.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
PAG_HOJAS = 1
F1 = 0
   
xl.Worksheets(1).Activate
Do Until nuc_repo.EOF ' loop 1
    If nuc_repo!CLI_SUBGRUPO <> WS_SUBGRUPO Then
'       If WS_SUBGRUPO <> "" Then GoSub TOTALES
'       If nuc_repo!CLI_SUBGRUPO = 50 Then
'          GoSub GENERAL
'       End If
       SQ_OPER = 1
       PUB_TIPREG = 333
       pu_zona = nuc_repo!CLI_SUBGRUPO
       PUB_NUMTAB = pu_zona
       PUB_CODCIA = LK_CODCIA
       LEER_TAB_LLAVE
       F1 = F1 + 1
       xl.Cells(F1, 1) = ""
       If tab_llave.EOF Then
          xl.Cells(F1, 2) = "......."
       Else
          xl.Cells(F1, 2) = Trim(tab_llave!tab_nomlargo)
       End If
       wranF = "B" & F1 & ":B" & F1
       xl.Range(wranF).Font.Bold = True
       xl.Range(wranF).Font.Size = 12
    End If

  WS_SUBGRUPO = nuc_repo!CLI_SUBGRUPO


    N = 0
    total_cob = 0
    total_ven = 0
    TOTAL_SALDO = 0
    TOTAL_VEN_DIA = 0
    FRMIMP.lblproceso.Caption = Trim(nuc_repo!cli_nombre)
    DoEvents
    pu_codcia = LK_CODCIA
   If nuc_repo!cli_codclie = 17735 Then
       Print ""
    End If

    TOTAL_SALDO = CAR_TOT_CPX("C", pu_codcia, nuc_repo!cli_codclie)
    If TOTAL_SALDO <= 0 Then GoTo otro_cli
    
    FRMIMP.ProgBar.Value = FRMIMP.ProgBar.Value + 1
    If nuc_repo!cli_codclie = 17735 Then
       Print ""
    End If
    
WS_SALDO1 = 0
WS_SALDO2 = 0
WS_SALDO3 = 0

WS_SALDO4 = 0
WW_SALDO1 = 0
WW_SALDO2 = 0
WW_SALDO3 = 0
WW_SALDO4 = 0
FX1 = 0
FX2 = 0
FX3 = 0
FX4 = 0
XX1 = 0
XX2 = 0
XX3 = 0
XX4 = 0

PSCAA(3) = WS_F1
GoSub BUSCA
WS_SALDO1 = WS_SALDO
If WS_SALDO <= 0 Then GoTo otro_cli

PSCAA(3) = WS_F2
GoSub BUSCA
WS_SALDO2 = WS_SALDO

PSCAA(3) = WS_F3
GoSub BUSCA
WS_SALDO3 = WS_SALDO

PSCAA(3) = WS_F4
GoSub BUSCA
WS_SALDO4 = WS_SALDO

PSCAA(3) = WW_F1
'GoSub BUSCA
WW_SALDO1 = WS_SALDO

PSCAA(3) = WW_F2
'GoSub BUSCA
WW_SALDO2 = WS_SALDO

PSCAA(3) = WW_F3
'GoSub BUSCA
WW_SALDO3 = WS_SALDO

PSCAA(3) = WW_F4
'GoSub BUSCA
WW_SALDO4 = WS_SALDO

Siguef:
   F1 = F1 + 1
   xcuenta = xcuenta + 1
   xl.Cells(F1, 1) = xcuenta
   xl.Cells(F1, 2) = Trim(nuc_repo!cli_nombre)
   
   xl.Cells(F1, 3) = WS_SALDO4
   xl.Cells(F1, 4) = WS_SALDO3
   xl.Cells(F1, 5) = WS_SALDO2
   xl.Cells(F1, 6) = WS_SALDO1
 
 '  xl.Cells(F1, 7) = WW_SALDO4
 '  xl.Cells(F1, 8) = WW_SALDO3
 '  xl.Cells(F1, 9) = WW_SALDO2
 '  xl.Cells(F1, 10) = WW_SALDO1
   
otro_cli:
   nuc_repo.MoveNext
Loop

Screen.MousePointer = 0
'GoSub TOTALES
F1 = F1 + 2
      
GoTo VAMOS

WEXCEL:
  Dim xlchart As Chart
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FRMIMP.lblproceso.Caption = "Abriendo , Archivo Cobranzas.xls . . . "
  DoEvents
  WPAS = "131296"
  xl.Workbooks.Open "C:\ADMIN\OFFICE\BALANCE.xls", 0, True, 4, WPAS, WPAS
Return

VAMOS:
  DoEvents
  FRMIMP.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  DoEvents
  
  
  xl.Visible = True
'  wranF = "F2:F4"
'  xl.Range(wranF).Borders.Item(xlEdgeLeft).LineStyle = 2
'  wranF = "J2:J4"
'  xl.Range(wranF).Borders.Item(xlEdgeLeft).LineStyle = 2
  wranF = "A5:F" & F1
  xl.Range(wranF).Borders.LineStyle = 2
  xl.Cells(1, 2) = "R: " & PUB_CODTRA
  xl.Worksheets("Hoja1").Range("A1:X51").Locked = True
  xl.Worksheets("Hoja1").Protect "131296"
  xl.Worksheets(1).Activate
  xl.Application.Visible = True
  DoEvents
  FRMIMP.lblproceso.Visible = False
  FRMIMP.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FRMIMP.Pantalla.Enabled = True
  FRMIMP.Pantalla.Caption = "Por &Pantalla"
  FRMIMP.lblproceso.Visible = False

Exit Sub



'*** RUTINAS PARA IMPRIMIR



WPROGRESO:

Return

Exit Sub
 



Exit Sub
CANCELA:
  FRMIMP.Pantalla.Enabled = True
  FRMIMP.Pantalla.Caption = "Por &Pantalla"
  FRMIMP.lblproceso.Visible = False
  LOC_CANCELA = 0
  Set xl = Nothing
  Screen.MousePointer = 0

Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FRMIMP

BUSCA:
    PSCAA(0) = LK_CODCIA
    PSCAA(1) = "C"
    PSCAA(2) = nuc_repo!cli_codclie
    CAA_REPO.Requery
    If CAA_REPO.EOF Then Return
    
CAA_REPO.MoveLast
N = 0
WS_SALDO = 0
Do Until CAA_REPO.BOF Or N = 1
   If CAA_REPO!CAA_IMPORTE > 0 And CAA_REPO!CAA_TIPDOC = "FA" And CAA_REPO!CAA_IMPORTE = CAA_REPO!CAA_TOTAL And CAA_REPO!CAA_FECHA = PSCAA(3) Then
   Else
      N = 1
      WS_SALDO = CAA_REPO!CAA_SALDO
   End If
   
   CAA_REPO.MovePrevious
Loop
Return

End Sub

Public Sub CLI_IMP()
' *** REPORTES DE NUCLEOS
On ERRO GoTo FINTODO
Dim success%
Dim LETRAS(24) As String * 1
Dim MESES(12) As String
Dim FECHA_INICIO
Dim WSFECHA As Date
Dim xMES1, xMES2 As Integer
Dim con  As Integer
Dim TABLAGEN() As String * 30
Dim TABLAGEN2() As String * 30
Dim TABPOLLOS() As Currency
Dim SUMPOLLOS() As Currency

Dim todos_cli As Integer
Dim filas_cli As Integer
Dim xpollos As String * 6
Dim xdias As String * 6
Dim C1 As Integer
Dim F1 As Integer
Dim WTABULA2 As String * 1
Dim WTABULA As String * 1
Dim RUTA As String
Dim WR_FECHA As String
Dim WR_LOTE As String
Dim WR_DAT1 As String
Dim SUM_cantidad As Currency
Dim WR_soles As String * 10
Dim WR_CIA As String
Dim cade As String

Dim SUM_SALDOS As Currency
Dim w_cant As String * 12
Dim w_sol As String * 10
Dim WCANT As Currency

Dim nuc_repo  As rdoResultset
Dim PSNUC_REPO As rdoQuery
Dim PSFAR_REPO As rdoQuery
Dim FAR_REPO  As rdoResultset
Dim PSLOT_REPO As rdoQuery
Dim lot_repo  As rdoResultset
Dim Mensaje, titulo, valorpred As String
Dim wvalor
Dim valor As Double
Dim xcuenta As Integer
Dim PAG_HOJAS As Integer
Dim WS_SALDO As Currency
Dim WS_S As String
Dim ws_tt As String

'FRMIMP.pantalla.Caption = "Cancelar"
'FRMIMP.pantalla.Enabled = True
DoEvents
cade = "SELECT * FROM CARACU WHERE CAA_CODCIA = ? AND CAA_CP = ? AND CAA_CODCLIE = ? AND CAA_FECHA >= ?  ORDER BY CAA_CODCLIE, CAA_FECHA, CAA_NUM_OPER"
Set PSCAA = CN.CreateQuery("", cade)
Set CAA_REPO = PSCAA.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

cade = "SELECT * FROM CLIENTES WHERE CLI_CODCIA = ? AND CLI_CP = ? AND CLI_ESTADO = 'B'  ORDER BY CLI_NOMBRE"
Set PSNUC_REPO = CN.CreateQuery("", cade)
Set nuc_repo = PSNUC_REPO.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

If wver = 99 Then FRMIMP.ProgBar.Value = 1

cade = "SELECT * FROM FACART WHERE FAR_CODCIA = ? AND FAR_CP = ? AND FAR_CODCLIE = ? AND FAR_TIPMOV = ?  ORDER BY FAR_CODCLIE, FAR_FECHA"
Set PSFAR_REPO = CN.CreateQuery("", cade)
Set FAR_REPO = PSFAR_REPO.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


'*** LLENA DE TABLAS ... CLIENTES NUCLEOS
FECHA_INICIO = LOC_FECHA_INC
PSNUC_REPO(0) = LK_CODCIA
PSNUC_REPO(1) = "C"
Screen.MousePointer = 11
SUM_SALDOS = 0
nuc_repo.Requery
If nuc_repo.EOF Then
  Screen.MousePointer = 0
  MsgBox "NO EXISTE NINGUNO ..", 48, Pub_Titulo
  FRMIMP.ProgBar.Visible = False
  FRMIMP.lblproceso.Visible = False
  Exit Sub
End If
FRMIMP.ProgBar.Min = 0
FRMIMP.ProgBar.Max = nuc_repo.RowCount
FRMIMP.ProgBar.Value = 0
FRMIMP.ProgBar.Visible = True
DoEvents
FRMIMP.lblproceso.Visible = True
FRMIMP.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
todos_cli = nuc_repo.RowCount
'todos_cli = 26
ReDim TABLAGEN(32, todos_cli * 2)
ReDim TABLAGEN2(32, todos_cli)
ReDim TABPOLLOS(32)
ReDim SUMPOLLOS(todos_cli)
Dim wFECHASIGUE As Date
Dim wtC, wtF

GoSub WEXCEL
PAG_HOJAS = 1

xl.Worksheets(1).Activate
GoSub LETRAS
'For C1 = 1 To 32
'For F1 = 1 To todos_cli
' TABLAGEN(C1, F1) = ""
'Next F1
'Next C1
WS_SALDO = 0
xcuenta = 0
F1 = 1
Do Until nuc_repo.EOF ' loop 1
  xcuenta = xcuenta + 1
  If LOC_CANCELA = 1 Then
    GoTo CANCELA
  End If
  FRMIMP.ProgBar.Value = xcuenta
  C1 = 0
  TABLAGEN(C1, F1) = Trim(nuc_repo!cli_nombre)
  SUM_cantidad = 0
  DoEvents
  FRMIMP.lblproceso.Caption = Trim(nuc_repo!cli_nombre) & " . . .  " & xcuenta & " de " & todos_cli
  DoEvents
  PSFAR_REPO(0) = LK_CODCIA
  PSFAR_REPO(1) = "C"
  PSFAR_REPO(2) = nuc_repo!cli_codclie
  PSFAR_REPO(3) = 28
  FAR_REPO.Requery
  If FAR_REPO.EOF Then
   ' GoTo Siguef
  End If
  QSALDO = 0
  QSALDO1 = 0
  QJ = 0
  
  wdia = 0
  WSFECHA = CDate(FECHA_INICIO)
  wFECHASIGUE = WSFECHA
  FRMIMP.ProgBar.Min = 0
  FRMIMP.ProgBar.Max = 13
  FRMIMP.ProgBar.Value = 0
  Do Until wdia >= 13
    FRMIMP.ProgBar.Value = wdia
    wFECHASIGUE = DateAdd("d", wdia, WSFECHA)
    If CDate(wFECHASIGUE) > CDate(LK_FECHA_DIA) Then
      Exit Do
    End If
    FAR_REPO.MoveFirst
    C1 = C1 + 1
    ws_total = 0
    ws_tt = ""
    TABLAGEN(C1, 0) = CDate(wFECHASIGUE)
    Do Until FAR_REPO.EOF
     If FAR_REPO!FAR_fecha = wFECHASIGUE Then
        ws_total = ws_total + FAR_REPO!FAR_CANTIDAD
     ElseIf FAR_REPO!FAR_fecha > wFECHASIGUE Then
        Exit Do
     End If
      FAR_REPO.MoveNext
    Loop
    FRMIMP.ProgBar.Value = FRMIMP.ProgBar.Max
    ws_tt = ws_total
    TABPOLLOS(wdia + 1) = TABPOLLOS(wdia + 1) + ws_total
    SUM_cantidad = SUM_cantidad + ws_total
    If ws_total = 0 Then ws_tt = ""
    xl.Cells(F1 + 7, C1 + 2) = "'" & ws_tt
    xl.Worksheets("Hoja1").Range(LETRAS(C1 + 2) & F1 + 7).HorizontalAlignment = xlRight
    xl.Worksheets("Hoja1").Range(LETRAS(C1 + 2) & F1 + 7 + 1).HorizontalAlignment = xlCenter
    wdia = wdia + 1
  Loop
  
    fx = 0
    PSCAA(0) = LK_CODCIA
    PSCAA(1) = "C"
    PSCAA(2) = nuc_repo!cli_codclie
    PSCAA(3) = CDate("7/11/98")
    CAA_REPO.Requery
    If CAA_REPO.EOF Then
       GoTo Siguef
    End If
    WDIF = DateDiff("d", WSFECHA, CAA_REPO!CAA_FECHA) + 3
    If WDIF > 3 Then
      QJ = 3
      LOC_VALOR = 0
      GoSub LLENA_VALOR
    End If
    QFECHA = CAA_REPO!CAA_FECHA
    QSALDO = CAA_REPO!CAA_SALDO
    QSALDO1 = 0
    QSALDO2 = 0
        
    Do Until CAA_REPO.EOF
    If QFECHA = CAA_REPO!CAA_FECHA Then
        GoTo SD
    Else
       WDIF = DateDiff("d", WSFECHA, CAA_REPO!CAA_FECHA) + 3
       
       If fx = 0 Then QSALDO = QSALDO2
       
       If WDIF >= 3 Then
         If QJ >= 3 Then xl.Cells(F1 + 7 + 1, QJ) = Format(QSALDO, "0.00")
               QJ = QJ + 1
               LOC_VALOR = QSALDO1
               GoSub LLENA_VALOR
       End If
        QFECHA = CAA_REPO!CAA_FECHA
        fx = 0
    End If
   QSALDO2 = QSALDO1
SD:
   
    QSALDO1 = CAA_REPO!CAA_SALDO

     If CAA_REPO!CAA_IMPORTE = CAA_REPO!caa_SALDO_car And (CAA_REPO!CAA_IMPORTE > 0 And CAA_REPO!CAA_TIPMOV > 0) Then
     Else
         fx = 1
         QSALDO = CAA_REPO!CAA_SALDO
     End If
     QJ = WDIF
     CAA_REPO.MoveNext
   Loop
   
   If fx = 0 Then QSALDO = QSALDO2
   WDIF = DateDiff("d", WSFECHA, LK_FECHA_DIA) + 3
   
   If WDIF >= QJ Then
   If WDIF >= 3 Then
        If QJ >= 3 Then xl.Cells(F1 + 7 + 1, QJ) = Format(QSALDO, "0.00")
           QJ = QJ + 1
           LOC_VALOR = QSALDO1
           GoSub LLENA_VALOR
  End If
 End If
Siguef:
 FRMIMP.ProgBar.Max = todos_cli
 SUMPOLLOS(xcuenta) = SUM_cantidad
 nuc_repo.MoveNext
 F1 = F1 + 2
Loop ' loop 1
FRMIMP.ProgBar.Value = todos_cli
Screen.MousePointer = 0
GoTo VAMOS

WEXCEL:
  Dim xlchart As Chart
  Dim wranF, wran1, wran2, WPAS
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FRMIMP.lblproceso.Caption = "Abriendo , Archivo Ventas.xls . . . "
  DoEvents
  WPAS = "131296"
  xl.Workbooks.Open "C:\ADMIN\OFFICE\ventas.xls", 0, True, 4, WPAS, WPAS
Return

VAMOS:
  DoEvents
  FRMIMP.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  DoEvents
  xcuenta = 1
  'XL.APPLICATION.VISIBLE = True
  For F1 = 1 To todos_cli * 2
    xl.Cells(F1 + 7, 2) = Trim(TABLAGEN(0, F1))
    'wranF = letras(17) & F1 + 7
    'wran1 = letras(3) & F1 + 7
    'wran2 = letras(16) & F1 + 7
    'xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
    
    If xcuenta <= (todos_cli * 2) Then
      xl.Cells(xcuenta + 7, 1) = F1
      xl.Cells(xcuenta + 7, 17) = Trim(SUMPOLLOS(F1))
    End If
    xcuenta = xcuenta + 2
  Next F1
  xcuenta = 0
  For F1 = 1 To todos_cli * 2
    If xcuenta <= (todos_cli * 2) Then
      xl.Cells(xcuenta + 7, 17) = " "
    End If
    xcuenta = xcuenta + 2
  Next F1
  
  WSFECHA = CDate(FECHA_INICIO)
  For C1 = 1 To 14
    wFECHASIGUE = DateAdd("d", C1 - 1, WSFECHA)
    xl.Cells(6, C1 + 2) = wFECHASIGUE
    xl.Cells(7, C1 + 2) = wFECHASIGUE
  Next C1
  
   xl.Cells((todos_cli * 2) + 8, 2) = " TOTALES - POLLOS"
   xl.Cells((todos_cli * 2) + 9, 2) = " TOTALES - SALDOS"
  xl.Range("Q7") = "TOTAL"
  
  wranF = "A" & (todos_cli * 2) + 8 & ":B" & (todos_cli * 2) + 9
  xl.Range(wranF).Font.Bold = True
 
  'wranF = "S8:S" & (todos_cli * 2) + 9
  'xl.Range(wranF).Font.Bold = True
  
  For F1 = 3 To 16
    wranF = LETRAS(F1) & (todos_cli * 2) + 9
    wran1 = LETRAS(F1) & 8
    wran2 = LETRAS(F1) & (todos_cli * 2) + 8 - 1
    xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
    wranF = LETRAS(F1) & (todos_cli * 2) + 8
    xl.Range(wranF) = Format(TABPOLLOS(F1 - 2), "#####")
  Next F1
  
  wran1 = "Q8"
  wran2 = "Q" & (todos_cli * 2) + 7
  wranF = "Q" & (todos_cli * 2) + 8
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  
  wranF = "C" & (todos_cli * 2) + 8 & ":Q" & (todos_cli * 2) + 9
  xl.Worksheets("Hoja1").Range(wranF).HorizontalAlignment = xlRight
  
  wranF = "A7:Q" & (todos_cli * 2) + 9
  xl.Range(wranF).Borders.LineStyle = 1

  WSFECHA = DateAdd("d", 15, CDate(FECHA_INICIO))
  xMES1 = Month(FECHA_INICIO)
  xMES2 = Month(WSFECHA)
  If xMES1 = xMES2 Then
    cade = Format(FECHA_INICIO, "mmmm")
  Else
    cade = Format(FECHA_INICIO, "mmmm") & " - " & Format(WSFECHA, "mmmm")
  End If
  xl.Cells(2, 2) = "R: " & PUB_CODTRA
  xl.Cells(3, 6) = "      R E P A R T O S    Y    S A L D O S  "
  'xl.Cells(3, 2) = Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption))
  xl.Worksheets("Hoja1").Range("A1:X51").Locked = True
  xl.Worksheets("Hoja1").Protect "131296"
  xl.Worksheets(1).Activate
  xl.Application.Visible = True
  DoEvents
  FRMIMP.lblproceso.Visible = False
  FRMIMP.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FRMIMP.Pantalla.Enabled = True
  FRMIMP.Pantalla.Caption = "Por &Pantalla"
  FRMIMP.lblproceso.Visible = False

Exit Sub



'*** RUTINAS PARA IMPRIMIR



WPROGRESO:

Return

Exit Sub
'SALE2:
'  MsgBox " Inesperado ..  " & Err.Description, 48, pub_titulo
 
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

LLENA_MESES:
MESES(1) = "ENERO"
MESES(1) = "FEBRERO"
MESES(1) = "MARZO"
MESES(1) = "ENERO"
MESES(1) = "ENERO"
MESES(1) = "ENERO"
MESES(1) = "ENERO"
MESES(1) = "ENERO"
MESES(1) = "ENERO"
MESES(1) = "ENERO"

Return

LLENA_VALOR:
For i = QJ To WDIF
  If i >= 3 Then
    xl.Cells(F1 + 7 + 1, i) = Format(LOC_VALOR, "0.00")
  End If
Next i
Return

Exit Sub
CANCELA:
  FRMIMP.Pantalla.Enabled = True
  FRMIMP.Pantalla.Caption = "Por &Pantalla"
  FRMIMP.lblproceso.Visible = False
  LOC_CANCELA = 0
  Set xl = Nothing
  Screen.MousePointer = 0

Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FRMIMP
 


End Sub
Public Function SALDO_CLIENTE(wCP As String, wcodcia As String, wcodclie As Double) As Double
Dim SUM_IMPORTE As Double
    SUM_IMPORTE = 0
    SQ_OPER = 2
    pu_codcia = wcodcia
    pu_codclie = wcodclie
    pu_cp = wCP
    LEER_CAR_LLAVE
    Do Until car_mayor.EOF
        If LK_FECHA_DIA = car_mayor!CAR_FECHA_INGR And car_mayor!CAR_IMPORTE > 0 Then
        Else
           SUM_IMPORTE = SUM_IMPORTE + (car_mayor!CAR_IMPORTE)
        End If
        car_mayor.MoveNext
    Loop
    SALDO_CLIENTE = SUM_IMPORTE
End Function

Public Sub PORPAGAR()
' *** REPORTES DE NUCLEOS
On ERRO GoTo FINTODO
Dim i As Integer
Dim fx
Dim nuc_repo  As rdoResultset
Dim PSNUC_REPO As rdoQuery
Dim PSCAR_REPO As rdoQuery
Dim CAR_REPO  As rdoResultset
Dim PARXX As rdoResultset
Dim TEMPOCIA(20) As String * 2
Dim TEMPOIMPORTE As Currency
Dim WT_TOT_IMPORTE_D As Currency
Dim WS_TOT_IMPORTE_DD As Currency
Dim Wt_TOT_IMPORTE_s As Currency
Dim WS_TOT_IMPORTE_ss As Currency
Dim tt_TOT_IMPORTE_s As Currency
Dim tt_TOT_IMPORTE_d As Currency

TEMPOCIA(1) = "01"
TEMPOCIA(2) = "03"
TEMPOCIA(3) = "02"
TEMPOCIA(4) = "04"

cade = "SELECT * FROM CLIENTES WHERE CLI_CP = ? AND CLI_CODCIA < '05' ORDER BY CLI_NOMBRE"
Set PSNUC_REPO = CN.CreateQuery("", cade)
Set nuc_repo = PSNUC_REPO.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

If wver = 99 Then FRMIMP.ProgBar.Value = 1

cade = "SELECT * FROM CARTERA WHERE  CAR_CP ='P' AND CAR_CODCLIE = ? AND CAR_CODCIA = ? AND CAR_IMPORTE <> 0 AND CAR_CODCIA > '00' AND CAR_CODCIA < '05' ORDER BY CAR_CODCLIE ,CAR_FECHA_VCTO"
Set PSCAR_REPO = CN.CreateQuery("", cade)
Set CAR_REPO = PSCAR_REPO.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

PSNUC_REPO(0) = "P"
nuc_repo.Requery
If nuc_repo.EOF Then
  Screen.MousePointer = 0
  MsgBox "NO EXISTE NINGUNO ..", 48, Pub_Titulo
  Exit Sub
End If


FRMIMP.ProgBar.Min = 0
FRMIMP.ProgBar.Max = nuc_repo.RowCount
FRMIMP.ProgBar.Value = 0
FRMIMP.ProgBar.Visible = True
DoEvents
FRMIMP.lblproceso.Visible = True
FRMIMP.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL

PAG_HOJAS = 1
xl.Worksheets(1).Activate
F1 = 5
wsNOMBRE = ""
xcuenta = 0
WS_TOT_IMPORTE = 0
Do Until nuc_repo.EOF ' loop 1
  'xcuenta = xcuenta + 1
  'FRMIMP.ProgBar.Value = xcuenta
  C1 = 0
  wentro = 0
  FRMIMP.lblproceso.Caption = Trim(nuc_repo!cli_nombre) & " . . .  " & xcuenta & " de " & todos_cli
  DoEvents
  PSCAR_REPO(0) = nuc_repo!cli_codclie
  PSCAR_REPO(1) = nuc_repo!CLI_CODCIA
  fx = 0
  CAR_REPO.Requery
  If CAR_REPO.EOF Then
     GoTo Siguef
  End If
If Trim(wsNOMBRE) <> Trim(nuc_repo!cli_nombre) Then
   If xcuenta > 0 Then
     GoSub SUBTOTALES
     F1 = F1 + 1
   Else
     F1 = F1 + 1
   End If
'   xl.Cells(F1, 13) = wsnombre
'   F1 = F1 + 1
     xcuenta = xcuenta + 1
     xl.Cells(F1, 1) = xcuenta & ".- " & nuc_repo!cli_nombre
     xl.Cells(F1, 13) = nuc_repo!cli_nombre
     xl.Cells(F1, 15) = -2500
     xl.Cells(F1, 1).Font.Bold = True
'     wranF = "A" & F1 & ":L" & F1
 '    xl.Range(wranF).Borders.Item(xlEdgeBottom).LineStyle = 1
     fila_activa = F1
     Wflag = "N"
     WT_TOT_IMPORTE_D = 0
     Wt_TOT_IMPORTE_s = 0
     WS_TOT_IMPORTE_DD = 0
     WS_TOT_IMPORTE_ss = 0
     RESTO = 0
End If
    wsNOMBRE = nuc_repo!cli_nombre
      WS_IMPORTE = 0
      Do Until CAR_REPO.EOF
        If nuc_repo!CLI_MONEDA = "D" Then
           WT_TOT_IMPORTE_D = WT_TOT_IMPORTE_D + CAR_REPO!CAR_IMPORTE
           tt_TOT_IMPORTE_d = tt_TOT_IMPORTE_d + CAR_REPO!CAR_IMPORTE
        Else
           Wt_TOT_IMPORTE_s = Wt_TOT_IMPORTE_s + CAR_REPO!CAR_IMPORTE
           tt_TOT_IMPORTE_s = tt_TOT_IMPORTE_s + CAR_REPO!CAR_IMPORTE
        End If
        
      
        If (LK_FECHA_DIA > CAR_REPO!car_fecha_vcto Or (nuc_repo!CLI_DET_TOT = "X" And fx = 0) Or (nuc_repo!CLI_DET_TOT <> "X" And nuc_repo!CLI_DET_TOT <> "Y")) And nuc_repo!CLI_DET_TOT <> "Y" Then
          GoSub imprime
        Else
        If nuc_repo!CLI_MONEDA = "D" Then
          WS_TOT_IMPORTE_DD = WS_TOT_IMPORTE_DD + CAR_REPO!CAR_IMPORTE
        Else
          WS_TOT_IMPORTE_ss = WS_TOT_IMPORTE_ss + CAR_REPO!CAR_IMPORTE
        End If
        If nuc_repo!CLI_DET_TOT = "X" Then RESTO = RESTO + 1
           
        ws_codcia = CAR_REPO!CAR_CODCIA
        End If
        
        fx = 1
        CAR_REPO.MoveNext
     Loop
WS_MONEDA = nuc_repo!CLI_MONEDA

Siguef:
nuc_repo.MoveNext
Loop
GoSub SUBTOTALES
F1 = F1 + 1
xl.Cells(F1, 1) = "              TOTAL GENERAL "
xl.Cells(F1, 3) = tt_TOT_IMPORTE_d
xl.Cells(F1, 4) = tt_TOT_IMPORTE_s
F1 = F1 + 1
xl.Cells(F1, 1) = "** FIN DE REPORTE **"
wranF = "A" & F1 & ":L" & F1
xl.Range(wranF).Borders.Item(xlEdgeBottom).LineStyle = 1


wranF = "A" & 6 & ":A" & F1
xl.Range(wranF).Borders.Item(xlEdgeRight).LineStyle = 1
wranF = "B" & 6 & ":B" & F1
xl.Range(wranF).Borders.Item(xlEdgeRight).LineStyle = 1
wranF = "D" & 6 & ":D" & F1
xl.Range(wranF).Borders.Item(xlEdgeRight).LineStyle = 1
wranF = "F" & 6 & ":F" & F1
xl.Range(wranF).Borders.Item(xlEdgeRight).LineStyle = 1
wranF = "H" & 6 & ":H" & F1
xl.Range(wranF).Borders.Item(xlEdgeRight).LineStyle = 1
wranF = "J" & 6 & ":J" & F1
xl.Range(wranF).Borders.Item(xlEdgeRight).LineStyle = 1
wranF = "L" & 6 & ":L" & F1
xl.Range(wranF).Borders.Item(xlEdgeRight).LineStyle = 1


'If F1 < 500 Then
'   wranF = "A" & F1 + 1 & ":L" & 500
'   xl.Range(wranF).Borders.LineStyle = 0
'   xl.Range(wranF).EntireRow.Select
'   xl.Range(wranF).EntireRow.Delete
'End If
'xl.Application.Visible = True
pub_mensaje = " Desea Ordenar por fecha ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbYes Then
   wranF = "A" & 6 & ":O" & F1
   xl.Application.Range(wranF).Select
   xl.Application.Range(wranF).Sort Key1:=Range("M1"), Order1:=xlAscending, Key2:=Range("O1"), Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

   fx = 1
   Do Until fx = F1
   If xl.Cells(fx, 15) = -2500 Then
      wranF = "A" & fx & ":L" & fx
      xl.Range(wranF).Borders.Item(xlEdgeBottom).LineStyle = 1
   End If
   xl.Cells(fx, 13) = ""
   xl.Cells(fx, 15) = ""
   fx = fx + 1
   Loop
End If

Screen.MousePointer = 0
  FRMIMP.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  DoEvents
  xl.Worksheets("Hoja1").Protect "131296"
  xl.Application.Visible = True
  DoEvents
  FRMIMP.lblproceso.Visible = False
  FRMIMP.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FRMIMP.Pantalla.Enabled = True
  FRMIMP.Pantalla.Caption = "Por &Pantalla"
  FRMIMP.lblproceso.Visible = False

Exit Sub

imprime:
TEMPOIMPORTE = CAR_REPO!CAR_IMPORTE
F1 = F1 + 1
xl.Cells(F1, 1) = "  " & Format(CAR_REPO!CAR_FECHA_INGR, "dd/mm/yy") & " " & Trim(CAR_REPO!car_concepto)
If CAR_REPO!car_fecha_vcto < #1/1/00# Then xl.Cells(F1, 2) = CAR_REPO!car_fecha_vcto
If CAR_REPO!car_fecha_vcto < #1/1/00# Then xl.Cells(F1, 15) = DateDiff("d", LK_FECHA_DIA, CAR_REPO!car_fecha_vcto)

If CAR_REPO!CAR_CODCIA = TEMPOCIA(1) Then
   If nuc_repo!CLI_MONEDA = "D" Then
     xl.Cells(F1, 5) = TEMPOIMPORTE
   Else
     xl.Cells(F1, 6) = TEMPOIMPORTE
   End If
ElseIf CAR_REPO!CAR_CODCIA = TEMPOCIA(2) Then
   If nuc_repo!CLI_MONEDA = "D" Then
     xl.Cells(F1, 7) = TEMPOIMPORTE
   Else
     xl.Cells(F1, 8) = TEMPOIMPORTE
   End If
ElseIf CAR_REPO!CAR_CODCIA = TEMPOCIA(3) Then
   If nuc_repo!CLI_MONEDA = "D" Then
     xl.Cells(F1, 9) = TEMPOIMPORTE
   Else
     xl.Cells(F1, 10) = TEMPOIMPORTE
   End If
ElseIf CAR_REPO!CAR_CODCIA = TEMPOCIA(4) Then
   If nuc_repo!CLI_MONEDA = "D" Then
     xl.Cells(F1, 11) = TEMPOIMPORTE
   Else
     xl.Cells(F1, 12) = TEMPOIMPORTE
   End If
End If
   If nuc_repo!CLI_MONEDA = "D" Then
      If TEMPOIMPORTE <> 0 Then xl.Cells(F1, 3) = TEMPOIMPORTE
     Else
      If TEMPOIMPORTE <> 0 Then xl.Cells(F1, 4) = TEMPOIMPORTE
   End If
xl.Cells(F1, 13) = nuc_repo!cli_nombre

Return

SUBTOTALES:
If WT_TOT_IMPORTE_D = WS_TOT_IMPORTE_DD And WS_TOT_IMPORTE_DD <> 0 Then
   TEMPOIMPORTE = WT_TOT_IMPORTE_D
   GoSub IMPRIMEY_d
   Return
ElseIf Wt_TOT_IMPORTE_s = WS_TOT_IMPORTE_ss And WS_TOT_IMPORTE_ss <> 0 Then
   TEMPOIMPORTE = Wt_TOT_IMPORTE_s
   GoSub IMPRIMEY_s
   Return
End If

F1 = F1 + 1
xl.Cells(F1, 13) = wsNOMBRE

If WS_TOT_IMPORTE_DD = 0 And WS_TOT_IMPORTE_ss = 0 Then
Else
If RESTO <> 0 Then
    xl.Cells(F1, 1) = "Saldo de " & RESTO & " Documentos"
Else
    xl.Cells(F1, 1) = "Saldo Sin Detalle "
End If
If WS_MONEDA = "D" Then
   xl.Cells(F1, 3) = WS_TOT_IMPORTE_DD
Else
   xl.Cells(F1, 4) = WS_TOT_IMPORTE_ss
End If
F1 = F1 + 1
xl.Cells(F1, 13) = wsNOMBRE
End If



If WT_TOT_IMPORTE_D <> 0 Then
   If WT_TOT_IMPORTE_D <> Val(xl.Cells(fila_activa + 1, 3)) Then xl.Cells(fila_activa, 3) = WT_TOT_IMPORTE_D
End If

If Wt_TOT_IMPORTE_s <> 0 Then
If Wt_TOT_IMPORTE_s <> Val(xl.Cells(fila_activa + 1, 4)) Then xl.Cells(fila_activa, 4) = Wt_TOT_IMPORTE_s
End If
Return


IMPRIMEY_s:
F1 = F1 + 1
xl.Cells(F1, 1) = "Saldo Por Guias "
If ws_codcia = TEMPOCIA(1) Then
     xl.Cells(F1, 6) = TEMPOIMPORTE
ElseIf ws_codcia = TEMPOCIA(2) Then
     xl.Cells(F1, 8) = TEMPOIMPORTE
ElseIf ws_codcia = TEMPOCIA(3) Then
     xl.Cells(F1, 10) = TEMPOIMPORTE
ElseIf ws_codcia = TEMPOCIA(4) Then
     xl.Cells(F1, 12) = TEMPOIMPORTE
End If
xl.Cells(F1, 4) = TEMPOIMPORTE
xl.Cells(F1, 13) = wsNOMBRE
F1 = F1 + 1
xl.Cells(F1, 13) = wsNOMBRE
Return

IMPRIMEY_d:
F1 = F1 + 1
xl.Cells(F1, 1) = "Saldo Por Guias "

If nuc_ws_CODCIA = TEMPOCIA(1) Then
     xl.Cells(F1, 5) = TEMPOIMPORTE
ElseIf ws_codcia = TEMPOCIA(2) Then
     xl.Cells(F1, 7) = TEMPOIMPORTE
ElseIf ws_codcia = TEMPOCIA(3) Then
     xl.Cells(F1, 9) = TEMPOIMPORTE
ElseIf ws_codcia = TEMPOCIA(4) Then
     xl.Cells(F1, 11) = TEMPOIMPORTE
End If
xl.Cells(F1, 3) = TEMPOIMPORTE
xl.Cells(F1, 13) = wsNOMBRE
F1 = F1 + 1
xl.Cells(F1, 13) = wsNOMBRE

Return

WEXCEL:
  Dim xlchart As Chart
  'Dim wranF, wran1, wran2, WPAS
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FRMIMP.lblproceso.Caption = "Abriendo , Archivo Ventas.xls . . . "
  DoEvents
  WPAS = "131296"
  xl.Workbooks.Open "C:\ADMIN\OFFICE\POR_PAGAR.xls", 0, True, 4, WPAS, WPAS
Return



'*** RUTINAS PARA IMPRIMIR



WPROGRESO:

Return

Exit Sub
CANCELA:
  FRMIMP.Pantalla.Enabled = True
  FRMIMP.Pantalla.Caption = "Por &Pantalla"
  FRMIMP.lblproceso.Visible = False
  LOC_CANCELA = 0
  Set xl = Nothing
  Screen.MousePointer = 0

Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FRMIMP

End Sub

Public Sub REPO_CAJA_GEN()
Dim ws_conta As Integer
Dim WS_NOMCLI As String * 20
Dim WS_NOMBAN As String * 20
FRMIMP.ProgBar.Value = 0
FRMIMP.ProgBar.Visible = True
DoEvents
FRMIMP.lblproceso.Visible = True
FRMIMP.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
PAG_HOJAS = 1

xl.Worksheets(1).Activate
WS_SALDO = 0
xcuenta = 0
F1 = 5

WS_MONEDA = txtNucleo.Text

If WS_MONEDA = "S" Then
   xl.Cells(4, 2) = "MONEDA:" & " NUEVOS SOLES"
ElseIf WS_MONEDA = "D" Then
   xl.Cells(4, 2) = "MONEDA:" & " DOLLARES"
End If


PUB_FECHA = LK_FECHA_DIA
SQ_OPER = 1
LEER_ALL_LLAVE
If all_llave.EOF Then GoTo VAMOS
FRMIMP.ProgBar.Min = 0
FRMIMP.ProgBar.Max = all_llave.RowCount

If WS_MONEDA = "S" Then
      WS_SALDO = par_llave!PAR_SALDO_CAJA_ayer
Else
      WS_SALDO = par_llave!PAR_SALDO_CAJA_D_ayer
End If
F1 = F1 + 1
xl.Cells(F1, 2) = "Saldo Anterior:"
xl.Cells(F1, 5) = WS_SALDO

ws_conta = 0
Do Until all_llave.EOF
   If all_llave!ALL_SIGNO_CAJA = 0 Then GoTo OTRO
   If all_llave!all_codCIA <> LK_CODCIA Then GoTo OTRO
   If all_llave!all_flag_ext = "E" Then GoTo OTRO
   If all_llave!ALL_SIGNO_CAR = 0 And all_llave!all_tipmov = 0 Then
      WS_IMPORTE = all_llave!ALL_IMPORTE
   Else
      WS_IMPORTE = all_llave!ALL_IMPORTE_AMORT
   End If
  
   If Trim(all_llave!all_moneda_ccm) <> " " And Val(all_llave!ALL_CODBAN) <> 0 Then
      ww_moneda = all_llave!all_moneda_ccm
   ElseIf Trim(all_llave!all_moneda_cli) <> " " And Val(all_llave!all_codclie) <> 0 Then
      ww_moneda = all_llave!all_moneda_cli
   ElseIf Trim(all_llave!ALL_MONEDA_CAJA) <> " " Then
      ww_moneda = all_llave!ALL_MONEDA_CAJA
   End If
   If ww_moneda <> WS_MONEDA Then GoTo OTRO
   WS_NOMCLI = ""
   WS_NOMBAN = ""
   ws_mensaje = ""
   WS_FBG = ""
   WS_LARGO = ""
   ws_nomche = ""
   WS_CODCLIE = Val(all_llave!all_codclie)
   If all_llave!all_numfac <> 0 Then
       If all_llave!ALL_FBG = "F" Then
          WS_FBG = "Fact. " & all_llave!all_numser & "-" & all_llave!all_numfac
       ElseIf all_llave!ALL_FBG = "B" Then
          WS_FBG = "Bolet." & all_llave!all_numser & "-" & all_llave!all_numfac
       Else
          WS_FBG = "Guia. " & all_llave!all_numser & "-" & all_llave!all_numfac
       End If
   End If
   If all_llave!ALL_CHENUM <> 0 And all_llave!ALL_SIGNO_CCM = -1 Then
       ws_nomche = "O/Pago: " & all_llave!ALL_CHENUM
   End If
      
   If WS_CODCLIE <> 0 Then
         pu_cp = all_llave!ALL_CP
         pu_codcia = LK_CODCIA
         pu_codclie = all_llave!all_codclie
         SQ_OPER = 1
         LEER_CLI_LLAVE
         WS_NOMCLI = Left(cli_llave!cli_nombre, 18) & ":"
   End If
   If Val(all_llave!ALL_CODBAN) <> 0 Then
         pu_codcia = LK_CODCIA
         PUB_CODBAN = all_llave!ALL_CODBAN
         SQ_OPER = 1
         LEER_CCM_LLAVE
         WS_NOMBAN = ccm_llave!CCM_nombre & ":"
   End If
   If all_llave!ALL_SIGNO_CAR <> 0 Then
       SQ_OPER = 1
       pu_cp = all_llave!ALL_CP
       pu_codclie = WS_CODCLIE
       pu_codcia = LK_CODCIA
       PUB_SERDOC = Nulo_Valor0(all_llave!ALL_serdoc)
       PUB_NUMDOC = all_llave!ALL_NUMDOC
       PUB_TIPDOC = all_llave!ALL_TIPDOC
       LEER_CAR_LLAVE
       WS_NOMCLI = Left(WS_NOMCLI, 30) & car_llave!CAR_FBG & "-." & car_llave!CAR_numSER & "-" & car_llave!CAR_numfac
   End If
   WS_LARGO = Trim(WS_NOMCLI) & Trim(WS_FBG) & Trim(ws_mensaje) & Trim(WS_NOMBAN) & Trim(ws_nomche)
   If WS_LARGO = "" Then WS_LARGO = all_llave!ALL_CONCEPTO
      
   F1 = F1 + 1
   If all_llave!ALL_SIGNO_CAJA = 1 Then
      WS_SALDO = WS_SALDO + WS_IMPORTE
      WS_SALDO_ING = WS_SALDO_ING + WS_IMPORTE
      xl.Cells(F1, 3) = WS_IMPORTE
   Else
      WS_SALDO = WS_SALDO - WS_IMPORTE
      WS_SALDO_SAL = WS_SALDO_SAL + WS_IMPORTE
      xl.Cells(F1, 4) = WS_IMPORTE
   End If
   ws_conta = ws_conta + 1
   xl.Cells(F1, 1) = ws_conta
   xl.Cells(F1, 2) = WS_LARGO
   xl.Cells(F1, 5) = WS_SALDO
   If Trim(all_llave!ALL_CONCEPTO) <> "" And Trim(WS_LARGO) <> Trim(all_llave!ALL_CONCEPTO) Then
      F1 = F1 + 1
      xl.Cells(F1, 2) = Trim(all_llave!ALL_CONCEPTO)
   End If
   
OTRO:
  all_llave.MoveNext
Loop
VAMOS:
par_llave.Edit
If WS_MONEDA = "S" Then
    par_llave!par_saldo_caja_hoy = WS_SALDO
Else
    par_llave!PAR_SALDO_CAJA_D_HOY = WS_SALDO
End If
par_llave.Update
wran1 = "C5"
wran2 = "C" & F1
xl.Visible = True
wranF = "C" & F1 + 1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
wran1 = "D5"
wran2 = "D" & F1
xl.Visible = True
wranF = "D" & F1 + 1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
wranF = "C" & 5 & ":C" & F1
xl.Range(wranF).Borders.Item(xlEdgeLeft).LineStyle = 3
wranF = "D" & 5 & ":D" & F1
xl.Range(wranF).Borders.Item(xlEdgeLeft).LineStyle = 3
wranF = "E" & 5 & ":E" & F1
xl.Range(wranF).Borders.Item(xlEdgeLeft).LineStyle = 3

FRMIMP.ProgBar.Value = todos_cli
Screen.MousePointer = 0
DoEvents
FRMIMP.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
DoEvents
xcuenta = 1


xl.Cells(2, 2) = "R: " & PUB_CODTRA
xl.Cells(2, 5) = LK_FECHA_DIA

xl.Worksheets("Hoja1").Range("A1:X51").Locked = True
xl.Worksheets("Hoja1").Protect "131296"
xl.Worksheets(1).Activate
xl.Application.Visible = True
DoEvents
FRMIMP.lblproceso.Visible = False
FRMIMP.ProgBar.Visible = False
Set xl = Nothing
Screen.MousePointer = 0
FRMIMP.Pantalla.Enabled = True
FRMIMP.Pantalla.Caption = "Por &Pantalla"
FRMIMP.lblproceso.Visible = False

Exit Sub


WEXCEL:
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FRMIMP.lblproceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
  DoEvents
  WPAS = "131296"
  xl.Workbooks.Open "C:\ADMIN\OFFICE\CAJA.xls", 0, True, 4, WPAS, WPAS
Return

Exit Sub

'LLENA_VALOR:
'For I = QJ To WDIF
'  If I >= 3 Then
'    xl.Cells(F1 + 7, I) = Format(LOC_VALOR, "0.00")
'  End If
'Next I
'Return

Exit Sub
CANCELA:
  FRMIMP.Pantalla.Enabled = True
  FRMIMP.Pantalla.Caption = "Por &Pantalla"
  FRMIMP.lblproceso.Visible = False
  LOC_CANCELA = 0
  Set xl = Nothing
  Screen.MousePointer = 0

Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FRMIMP
Exit Sub

End Sub
Public Sub REPO_CAJA_CHICA()
Dim ws_conta As Integer
FRMIMP.ProgBar.Value = 0
FRMIMP.ProgBar.Visible = True
DoEvents
FRMIMP.lblproceso.Visible = True
FRMIMP.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
PAG_HOJAS = 1

xl.Worksheets(1).Activate
WS_SALDO = 0
xcuenta = 0
F1 = 5


fila = 0
PUB_FECHA = DateAdd("d", -30, LK_FECHA_DIA)
PUB_CODBAN = 1
SQ_OPER = 4
LEER_CHE_LLAVE
If che_repo.EOF = True Then
   Screen.MousePointer = 11
   MsgBox "No hay Movimientos..."
   GoTo VAMOS
End If
ws_conta = 0
che_repo.MoveLast
Do Until che_repo.BOF
   ws_conta = ws_conta + 1
   If che_repo!CHE_SIGNO_CCM = 1 And che_repo!che_importe > 0 And che_repo!che_estado <> "E" Then Exit Do
   che_repo.MovePrevious
Loop

FRMIMP.ProgBar.Min = 0
FRMIMP.ProgBar.Max = ws_conta

WS_SALDO = che_repo!che_saldo

F1 = F1 + 1
xl.Cells(F1, 2) = "Saldo Anterior:"
xl.Cells(F1, 9) = WS_SALDO
      
ws_conta = 0
ws_fecha = che_repo!CHE_FECHA
wranF = "A" & F1 + 1 & ":I" & F1 + 1
xl.Range(wranF).Font.Bold = True
ws_saldo_dia = 0
Do Until che_repo.EOF
   If che_repo!CHE_FECHA <> ws_fecha Then
      xl.Cells(F1, 8) = ws_saldo_dia
      F1 = F1 + 1
      ws_saldo_dia = 0
   End If
   
   If che_repo!che_estado = "E" Then GoTo OTRO
   If F1 > 6 And fx = 0 Then
    wranF = "A" & F1 & ":I" & F1
    xl.Worksheets("Hoja1").Range(wranF).Borders.LineStyle = 2
    fx = 1
   End If
   
   
   
   ws_conta = ws_conta + 1
   FRMIMP.ProgBar.Value = ws_conta
   DoEvents
   F1 = F1 + 1
   WS_SALDO = Val(WS_SALDO) + Val(che_repo!che_importe) * che_repo!CHE_SIGNO_CCM
   If che_repo!CHE_SIGNO_CCM < 0 Then ws_saldo_dia = Val(ws_saldo_dia) + Val(che_repo!che_importe)
   xl.Cells(F1, 1) = che_repo!CHE_FECHA
   xl.Cells(F1, 2) = che_repo!che_concepto
   xl.Cells(F1, 7) = che_repo!che_importe
   xl.Cells(F1, 9) = WS_SALDO
   ws_fecha = che_repo!CHE_FECHA
   
OTRO:
  che_repo.MoveNext
Loop
      xl.Cells(F1, 8) = ws_saldo_dia
      F1 = F1 + 1
      xl.Cells(F1, 9) = WS_SALDO
      xl.Cells(F1, 6) = "Saldo en Caja Chica"
      wranF = "E" & F1 & ":I" & F1
      xl.Range(wranF).Font.Bold = True
  


Screen.MousePointer = 0
GoTo VAMOS

WEXCEL:
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FRMIMP.lblproceso.Caption = "Abriendo , Archivo Caja_Chica.xls . . . "
  DoEvents
  WPAS = "131296"
  xl.Workbooks.Open "C:\ADMIN\OFFICE\CAJA_CHICA.xls", 0, True, 4, WPAS, WPAS
Return

VAMOS:
  DoEvents
  FRMIMP.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  DoEvents
  xcuenta = 1
  
  
  xl.Cells(2, 2) = "R: " & PUB_CODTRA
  xl.Cells(2, 9) = LK_FECHA_DIA
  
  xl.Worksheets("Hoja1").Range("A1:X51").Locked = True
  xl.Worksheets("Hoja1").Protect "131296"
  xl.Worksheets(1).Activate
  xl.Application.Visible = True
  DoEvents
  FRMIMP.lblproceso.Visible = False
  FRMIMP.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FRMIMP.Pantalla.Enabled = True
  FRMIMP.Pantalla.Caption = "Por &Pantalla"
  FRMIMP.lblproceso.Visible = False

Exit Sub



'*** RUTINAS PARA IMPRIMIR



WPROGRESO:

Return

Exit Sub
'SALE2:
'  MsgBox " Inesperado ..  " & Err.Description, 48, pub_titulo
 
Return

LLENA_MESES:

Return

LLENA_VALOR:
For i = QJ To WDIF
  If i >= 3 Then
    xl.Cells(F1 + 7, i) = Format(LOC_VALOR, "0.00")
  End If
Next i
Return

Exit Sub
CANCELA:
  FRMIMP.Pantalla.Enabled = True
  FRMIMP.Pantalla.Caption = "Por &Pantalla"
  FRMIMP.lblproceso.Visible = False
  LOC_CANCELA = 0
  Set xl = Nothing
  Screen.MousePointer = 0

Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FRMIMP
Exit Sub






End Sub

Public Sub RECORD_VENDEDOR()
' *** REPORTES DE NUCLEOS
On ERRO GoTo FINTODO
Dim far_VEN  As rdoResultset
Dim PSFAR_VEN As rdoQuery
Dim tab_ventas(12) As Currency
Dim WS_NETO


pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA >= ? AND  FAR_ESTADO <> 'E' and far_tipmov  =10 ORDER BY FAR_CODVEN, FAR_FECHA,far_fbg, far_numser, far_numfac "
Set PSFAR_VEN = CN.CreateQuery("", pub_cadena)
Set far_VEN = PSFAR_VEN.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PSFAR_VEN(0) = LK_CODCIA
PSFAR_VEN(1) = DateAdd("y", -1, LK_FECHA_DIA)
far_VEN.Requery

If far_VEN.EOF Then
  Screen.MousePointer = 0
  MsgBox "NO EXISTE NINGUNO ..", 48, Pub_Titulo
  Exit Sub
End If
GoSub WEXCEL

i = 1
Do Until i = 6
   xl.Cells(3, i + 1) = DateAdd("m", -i + 1, LK_FECHA_DIA)
   i = i + 1
Loop


'FRMIMP.ProgBar.Value = FRMIMP.ProgBar.Value + 1
DoEvents
F1 = 5
WS_codven = 0
   WS_FBG = far_VEN!far_fbg
   WS_NUMSER = far_VEN!far_numser
   WS_NUMFAC = far_VEN!far_numfac
   WS_codven = far_VEN!FAR_codven
   WS_BRUTO = far_VEN!far_bruto
   WS_DESCTO = far_VEN!FAR_TOT_DESCTO
   WS_NETO = Val(far_VEN!far_bruto) - far_VEN!FAR_TOT_DESCTO + far_VEN!far_impto
Do Until far_VEN.EOF
   If WS_codven <> far_VEN!FAR_codven Then
      GoSub imprime
   End If
   If WS_FBG = far_VEN!far_fbg And WS_NUMSER = far_VEN!far_numser And WS_NUMFAC = far_VEN!far_numfac Then
   Else
      If ws_fecha > xl.Cells(3, 2) Then
         WS_VENTAS1 = WS_VENTAS1 + WS_NETO
      ElseIf ws_fecha > xl.Cells(3, 3) Then
         WS_VENTAS2 = WS_VENTAS2 + WS_NETO
      ElseIf ws_fecha > xl.Cells(3, 4) Then
         WS_VENTAS3 = WS_VENTAS3 + WS_NETO
      ElseIf ws_fecha > xl.Cells(3, 5) Then
         WS_VENTAS4 = WS_VENTAS4 + WS_NETO
      ElseIf ws_fecha > xl.Cells(3, 6) Then
         WS_VENTAS5 = WS_VENTAS5 + WS_NETO
      ElseIf ws_fecha > xl.Cells(3, 7) Then
         WS_VENTAS6 = WS_VENTAS6 + WS_NETO
      End If
   End If
   WS_FBG = far_VEN!far_fbg
   WS_NUMSER = far_VEN!far_numser
   WS_NUMFAC = far_VEN!far_numfac
   WS_codven = far_VEN!FAR_codven
   ws_fecha = far_VEN!FAR_fecha
   WS_NETO = Val(far_VEN!far_bruto) - far_VEN!FAR_TOT_DESCTO + far_VEN!far_impto
   far_VEN.MoveNext
Loop

GoSub imprime

F1 = 2
FRMIMP.lblproceso.Caption = "Cargando Datos al Excel . . . "
wranF = "B" & F1 & ":F" & F1
xl.Range(wranF).Font.Bold = True
xl.Cells(2, 2) = "REPORTE DE VENTAS X VENDEDORES "

DoEvents
PAG_HOJAS = 1
fila = 1
F1 = 6
xl.Worksheets(1).Activate
FRMIMP.ProgBar.Value = 0

  
Screen.MousePointer = 0
FRMIMP.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
FRMIMP.ProgBar.Value = FRMIMP.ProgBar.Max
DoEvents
xl.Visible = True
DoEvents
FRMIMP.lblproceso.Visible = False
FRMIMP.ProgBar.Visible = False
Set xl = Nothing
Screen.MousePointer = 0
FRMIMP.Pantalla.Enabled = True
FRMIMP.Pantalla.Caption = "Por &Pantalla"
FRMIMP.lblproceso.Visible = False

Exit Sub
imprime:
      SQ_OPER = 1
      PUB_CODVEN = WS_codven
      pu_codcia = LK_CODCIA
      LEER_VEN_LLAVE
       xl.Cells(F1, 1) = ven_llave!VEM_NOMBRE
       xl.Cells(F1, 2) = WS_VENTAS1
       xl.Cells(F1, 3) = WS_VENTAS2
       xl.Cells(F1, 4) = WS_VENTAS3
       xl.Cells(F1, 5) = WS_VENTAS4
       xl.Cells(F1, 6) = WS_VENTAS5
       xl.Cells(F1, 7) = WS_VENTAS6
       WS_VENTAS1 = 0
       WS_VENTAS2 = 0
       WS_VENTAS3 = 0
       WS_VENTAS4 = 0
       WS_VENTAS5 = 0
       WS_VENTAS6 = 0
       
Return


WEXCEL:
  Dim xlchart As Chart
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FRMIMP.lblproceso.Caption = "Abriendo , Archivo CONVERSION.xls . . . "
  DoEvents
  WPAS = "131296"
  xl.Workbooks.Open "C:\ADMIN\OFFICE\VENDEDORES.xls", 0, True, 4, WPAS, WPAS
Return
CANCELA:
  FRMIMP.Pantalla.Enabled = True
  FRMIMP.Pantalla.Caption = "Por &Pantalla"
  FRMIMP.lblproceso.Visible = False
  LOC_CANCELA = 0
  Set xl = Nothing
  Screen.MousePointer = 0

Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FRMIMP

End Sub


Public Sub IMP_FACTURACION()
On ERRO GoTo FINTODO
Dim wranF, wran1, wran2, WPAS
Dim LETRAS(24) As String * 1
Dim FECHA_INICIO
Dim WSFECHA As Date
Dim TABLAGEN2() As String * 30
Dim i, WS_SUBGRUPO
Dim C1 As Integer
Dim F1 As Integer
Dim nuc_repo  As rdoResultset
Dim PSNUC_REPO As rdoQuery
Dim far_r  As rdoResultset
Dim PSFAR As rdoQuery
Dim Mensaje, titulo, valorpred As String
Dim wvalor
Dim valor As Double
Dim xcuenta As Integer
Dim PAG_HOJAS As Integer
Dim WS_BRUTO As Currency
Dim w_fbg As String * 1
Dim w_serie As Integer
Dim w_numfac_ini As Currency
Dim w_numfac_fin As Currency
Dim llave_rep03 As rdoResultset
Dim PS_REP03 As rdoQuery


cade = "SELECT * FROM CARACU WHERE CAA_CODCIA = ? AND CAA_CP = ? AND CAA_CODCLIE = ? AND CAA_FECHA >= ?  ORDER BY CAA_CODCLIE, CAA_FECHA, CAA_NUM_OPER"
Set PSCAA = CN.CreateQuery("", cade)
Set CAA_REPO = PSCAA.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
DoEvents
cade = "SELECT * FROM CLIENTES WHERE CLI_CODCIA = ? AND CLI_CP = ? AND CLI_ESTADO = 'B'  ORDER BY CLI_SUBGRUPO, CLI_NOMBRE"
Set PSNUC_REPO = CN.CreateQuery("", cade)
Set nuc_repo = PSNUC_REPO.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


pub_cadena = "SELECT * FROM facart WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_NUMSER = ? AND FAR_FBG=? AND FAR_NUMFAC >= ? AND FAR_NUMFAC <= ? AND FAR_ESTADO <> 'E' ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_NUMSER, FAR_FBG, FAR_NUMFAC, FAR_NUMSEC"
Set PSFAR = CN.CreateQuery("", pub_cadena)
Set far_r = PSFAR.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

If wver = 99 Then FRMIMP.ProgBar.Value = 1
w_serie = Val(txtSerie.Text)
w_fbg = Left(CMBFGB.Text, 1)
w_numfac_ini = Val(txtnumfac1.Text)
w_numfac_fin = Val(txtnumfac2.Text)

'*** VERFICA SI HAY DATOS , O ESTAN CORRECTOS

PSFAR.rdoParameters(0) = 10
PSFAR.rdoParameters(1) = LK_CODCIA
PSFAR.rdoParameters(2) = w_serie
PSFAR.rdoParameters(3) = w_fbg
PSFAR.rdoParameters(4) = w_numfac_ini
PSFAR.rdoParameters(5) = w_numfac_fin
far_r.Requery
If far_r.EOF = True Then
   MsgBox "!!! NO EXISTEN Documentos ...", 48, Pub_Titulo
   GoTo CANCELA
End If
If far_r!FAR_CODCIA <> LK_CODCIA Then
   MsgBox "!!! NO TE CORRESPONDE...", 48, Pub_Titulo
   GoTo CANCELA
End If
If far_r!FAR_fecha <> LK_FECHA_DIA Then
'   MsgBox "!!! NO ES DEL DIA...", 48, Pub_Titulo
'   GoTo CANCELA
End If
ws_clave = "0"
usu.Requery
Do Until usu.EOF
  If Trim(usu!usu_key) = "ADMIN" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  usu.MoveNext
Loop

FRMIMP.ProgBar.Min = 0
FRMIMP.ProgBar.Max = far_r.RowCount
FRMIMP.ProgBar.Value = 0
FRMIMP.ProgBar.Visible = True
DoEvents
FRMIMP.lblproceso.Visible = True
FRMIMP.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents

GoSub WEXCEL
If Left(CMBFGB.Text, 1) = "B" Then
   xl.Worksheets(1).Activate
ElseIf Left(CMBFGB.Text, 1) = "F" Then
   xl.Worksheets(2).Activate
ElseIf Left(CMBFGB.Text, 1) = "G" Then
   xl.Worksheets(3).Activate
End If
GoSub LETRAS
wmaxitem = 15
waumenta = 0
WS_SALDO = 0
xcuenta = 0
F1 = 0
FILA_INICIAL = 0
wnumfac = far_r!far_numfac
fila = 0
WS_BRUTO = 0
SUB_CANT = 0
subtotal = 0
PUB_DESCTO = 0
Do Until far_r.EOF
   FRMIMP.ProgBar.Value = FRMIMP.ProgBar.Value + 1
   If far_r!FAR_NUMSEC = 1 Then
   F1 = F1 + 1
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   pu_cp = "C"
   pu_codclie = far_r!far_codclie
   LEER_CLI_LLAVE
   xl.Cells(F1, 2) = Trim(cli_llave!cli_nombre)
   xl.Cells(F1, 3) = far_r!far_numfac
   xl.Cells(F1, 4) = far_r!far_bruto - far_r!FAR_TOT_DESCTO
   End If
   
   far_r.MoveNext
Loop

  xl.Application.Visible = True
  DoEvents
  FRMIMP.lblproceso.Visible = False
  FRMIMP.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FRMIMP.Pantalla.Enabled = True
  FRMIMP.Pantalla.Caption = "Por &Pantalla"
  FRMIMP.lblproceso.Visible = False

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

  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FRMIMP.lblproceso.Caption = "Abriendo , Archivo Facturación.xls . . . "
  DoEvents
  WPAS = "131296"
  xl.Workbooks.Open "C:\ADMIN\OFFICE\Facturacion.xls", 0, True, 4, WPAS, WPAS
Return

Exit Sub
CANCELA:
  FRMIMP.Pantalla.Enabled = True
  FRMIMP.Pantalla.Caption = "Por &Pantalla"
  FRMIMP.lblproceso.Visible = False
  LOC_CANCELA = 0
  Set xl = Nothing
  Screen.MousePointer = 0

Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FRMIMP
 
End Sub


Private Sub txtnumfac1_KeyPress(KeyAscii As Integer)
ENTERO_VALOR KeyAscii
If KeyAscii = 13 Then
  txtnumfac2.SetFocus
End If

End Sub

Private Sub txtnumfac2_KeyPress(KeyAscii As Integer)
ENTERO_VALOR KeyAscii
If KeyAscii = 13 Then
  Pantalla.SetFocus
End If

End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
ENTERO_VALOR KeyAscii
If KeyAscii = 13 Then
  txtnumfac1.SetFocus
End If


End Sub
Public Sub ENTERO_VALOR(Optional tecla)
'CONVIERTE TODA A MAYUSCULAS LETRAS
Dim car As String, Longt As Integer
car = Chr$(tecla)
car = UCase$(Chr$(tecla))
tecla = Asc(car)
If car < "0" Or car > "9" Then
    If tecla <> 8 And tecla <> 13 And tecla <> 32 Then
        tecla = 0
        Beep
    End If
End If
End Sub
Public Sub IMP_STOCK()
On ERRO GoTo FINTODO
Dim wranF, wran1, wran2, WPAS
Dim LETRAS(24) As String * 1
Dim FECHA_INICIO
Dim WSFECHA As Date
Dim TABLAGEN2() As String * 30
Dim i, WS_SUBGRUPO
Dim C1 As Integer
Dim F1 As Integer
Dim nuc_repo  As rdoResultset
Dim PSNUC_REPO As rdoQuery
Dim far_r  As rdoResultset
Dim PSFAR As rdoQuery
Dim Mensaje, titulo, valorpred As String
Dim wvalor
Dim valor As Double
Dim xcuenta As Integer
Dim PAG_HOJAS As Integer
Dim WS_BRUTO As Currency
Dim w_fbg As String * 1
Dim w_serie As Integer
Dim w_numfac_ini As Currency
Dim w_numfac_fin As Currency


cade = "SELECT * FROM CARACU WHERE CAA_CODCIA = ? AND CAA_CP = ? AND CAA_CODCLIE = ? AND CAA_FECHA >= ?  ORDER BY CAA_CODCLIE, CAA_FECHA, CAA_NUM_OPER"
Set PSCAA = CN.CreateQuery("", cade)
Set CAA_REPO = PSCAA.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
DoEvents
cade = "SELECT * FROM CLIENTES WHERE CLI_CODCIA = ? AND CLI_CP = ? AND CLI_ESTADO = 'B'  ORDER BY CLI_SUBGRUPO, CLI_NOMBRE"
Set PSNUC_REPO = CN.CreateQuery("", cade)
Set nuc_repo = PSNUC_REPO.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


pub_cadena = "SELECT * FROM facart WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_NUMSER = ? AND FAR_FBG=? AND FAR_NUMFAC >= ? AND FAR_NUMFAC <= ? AND FAR_ESTADO <> 'E' ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_NUMSER, FAR_FBG, FAR_NUMFAC, FAR_NUMSEC"
Set PSFAR = CN.CreateQuery("", pub_cadena)
Set far_r = PSFAR.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

If wver = 99 Then FRMIMP.ProgBar.Value = 1
w_serie = Val(txtSerie.Text)
w_fbg = Left(CMBFGB.Text, 1)
w_numfac_ini = Val(txtnumfac1.Text)
w_numfac_fin = Val(txtnumfac2.Text)

'*** VERFICA SI HAY DATOS , O ESTAN CORRECTOS

PSFAR.rdoParameters(0) = 10
PSFAR.rdoParameters(1) = LK_CODCIA
PSFAR.rdoParameters(2) = w_serie
PSFAR.rdoParameters(3) = w_fbg
PSFAR.rdoParameters(4) = w_numfac_ini
PSFAR.rdoParameters(5) = w_numfac_fin
far_r.Requery
If far_r.EOF = True Then
   MsgBox "!!! NO EXISTEN Documentos ...", 48, Pub_Titulo
   GoTo CANCELA
End If
If far_r!FAR_CODCIA <> LK_CODCIA Then
   MsgBox "!!! NO TE CORRESPONDE...", 48, Pub_Titulo
   GoTo CANCELA
End If
If far_r!FAR_fecha <> LK_FECHA_DIA Then
'   MsgBox "!!! NO ES DEL DIA...", 48, Pub_Titulo
'   GoTo CANCELA
End If
ws_clave = "0"
usu.Requery
Do Until usu.EOF
  If Trim(usu!usu_key) = "ADMIN" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  usu.MoveNext
Loop

FRMIMP.ProgBar.Min = 0
FRMIMP.ProgBar.Max = far_r.RowCount
FRMIMP.ProgBar.Value = 0
FRMIMP.ProgBar.Visible = True
DoEvents
FRMIMP.lblproceso.Visible = True
FRMIMP.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents

GoSub WEXCEL
If Left(CMBFGB.Text, 1) = "B" Then
   xl.Worksheets(1).Activate
ElseIf Left(CMBFGB.Text, 1) = "F" Then
   xl.Worksheets(2).Activate
ElseIf Left(CMBFGB.Text, 1) = "G" Then
   xl.Worksheets(3).Activate
End If
GoSub LETRAS
wmaxitem = 15
waumenta = 0
WS_SALDO = 0
xcuenta = 0
F1 = 0
FILA_INICIAL = 0
wnumfac = far_r!far_numfac
F1 = F1 + 1
xl.Cells(F1, 9) = "'" & far_r!far_numser & " - " & far_r!far_numfac
F1 = F1 + 1
SQ_OPER = 1
pu_codcia = LK_CODCIA
pu_cp = "C"
'xl.Application.Visible = True
pu_codclie = far_r!far_codclie
LEER_CLI_LLAVE
If cli_llave.EOF Then
  MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
  GoTo CANCELA
End If
xl.Cells(F1, 2) = Trim(cli_llave!cli_nombre)
xl.Cells(F1, 7) = "'" & Format(far_r!FAR_fecha, "dd/mm/yyyy")
F1 = F1 + 1
xl.Cells(F1, 2) = Trim(cli_llave!CLI_CASA_DIREC) & " # " & cli_llave!CLI_CASA_NUM
If Left(CMBFGB.Text, 1) = "F" Then
  xl.Cells(F1, 7) = "'" & Trim(cli_llave!cli_ruc_esposo)
End If
F1 = F1 + 2
fila = 0
WS_BRUTO = 0
SUB_CANT = 0
subtotal = 0
PUB_DESCTO = 0
Do Until far_r.EOF
   FRMIMP.ProgBar.Value = FRMIMP.ProgBar.Value + 1
   If wnumfac <> far_r!far_numfac Then
'   xl.Application.Visible = True
    waumenta = wmaxitem - fila
    F1 = F1 + waumenta + 1
    xl.Cells(F1 - 1, 2) = "SON: " & CONVER_LETRAS(WS_BRUTO, "S")
    'If Left(CMBFGB.text, 1) = "F" Then
      xl.Cells(F1, 2) = WS_BRUTO - WS_IMPTO + WS_DESCTO
      xl.Cells(F1, 3) = WS_DESCTO
      xl.Cells(F1, 4) = WS_IMPTO
    'End If
    'If Left(CMBFGB.text, 1) = "F" Or Left(CMBFGB.text, 1) = "B" Then
      xl.Cells(F1, 8) = WS_BRUTO
    'End If
    F1 = F1 + 1
    xl.Cells(F1, 9) = "'" & far_r!far_numser & " - " & far_r!far_numfac
    F1 = F1 + 1
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    pu_cp = "C"
    pu_codclie = far_r!far_codclie
    LEER_CLI_LLAVE
    If cli_llave.EOF Then
      MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
      GoTo CANCELA
    End If
    xl.Cells(F1, 2) = Trim(cli_llave!cli_nombre)
    xl.Cells(F1, 7) = "'" & Format(far_r!FAR_fecha, "dd/mm/yyyy")
    F1 = F1 + 1
    xl.Cells(F1, 2) = Trim(cli_llave!CLI_CASA_DIREC) & " # " & cli_llave!CLI_CASA_NUM
    If Left(CMBFGB.Text, 1) = "F" Then
     xl.Cells(F1, 7) = "'" & Trim(cli_llave!cli_ruc_esposo)
    End If
    wnumfac = far_r!far_numfac
    'xl.Application.Visible = True
    F1 = F1 + 2
    fila = 0
    WS_BRUTO = 0
    SUB_CANT = 0
    subtotal = 0
    PUB_DESCTO = 0
   End If
   fila = fila + 1
   PUB_KEY = far_r!far_codart
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_ART_LLAVE
   If art_LLAVE.EOF Then
      MsgBox "Error Grave en arti..."
      GoTo CANCELA
   Else
      xl.Cells(F1, 2) = Trim(art_LLAVE!ART_NOMBRE)
   End If
   'xl.Cells(F1, 5) = far_r!far_unidades
   xl.Cells(F1, 6) = far_r!FAR_CANTIDAD
   If Left(CMBFGB.Text, 1) <> "G" Then
    xl.Cells(F1, 7) = far_r!far_PRECIO
   End If
   subtotal = Format(far_r!far_PRECIO * far_r!FAR_CANTIDAD)
   If far_r!FAR_DESCTO <> 0 Then
     subtotal = 0
   End If
   subtotal = redondea(subtotal)
   'If Left(CMBFGB.text, 1) <> "G" Then
    xl.Cells(F1, 8) = subtotal
   'End If
   WS_BRUTO = WS_BRUTO + subtotal
   SUB_CANT = SUB_CANT + far_r!FAR_CANTIDAD
   WS_DESCTO = far_r!FAR_TOT_DESCTO
   WS_IMPTO = far_r!far_impto
   WS_GASTOS = far_r!FAR_GASTOS
   F1 = F1 + 1
   far_r.MoveNext
Loop
  waumenta = wmaxitem - fila
  F1 = F1 + waumenta + 1
  xl.Cells(F1 - 1, 2) = "SON: " & CONVER_LETRAS(WS_BRUTO, "S")
  'If Left(CMBFGB.text, 1) = "F" Then
   xl.Cells(F1, 2) = WS_BRUTO - WS_IMPTO + WS_DESCTO
   xl.Cells(F1, 3) = WS_DESCTO
   xl.Cells(F1, 4) = WS_IMPTO
  'End If
  'If Left(CMBFGB.text, 1) = "F" Or Left(CMBFGB.text, 1) = "B" Then
   xl.Cells(F1, 8) = WS_BRUTO
  'End If
  DoEvents
  FRMIMP.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  DoEvents
  
  'wranF = "B" & F1 & ":B" & F1
  'xl.Range(wranF).Font.Bold = True
  If Left(CMBFGB.Text, 1) = "B" Then
    xl.Worksheets(1).Protect ws_clave
  ElseIf Left(CMBFGB.Text, 1) = "F" Then
    xl.Worksheets(2).Protect ws_clave
  ElseIf Left(CMBFGB.Text, 1) = "G" Then
    xl.Worksheets(3).Protect ws_clave
  End If

  xl.Application.Visible = True
  DoEvents
  FRMIMP.lblproceso.Visible = False
  FRMIMP.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FRMIMP.Pantalla.Enabled = True
  FRMIMP.Pantalla.Caption = "Por &Pantalla"
  FRMIMP.lblproceso.Visible = False

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

  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FRMIMP.lblproceso.Caption = "Abriendo , Archivo Facturación.xls . . . "
  DoEvents
  WPAS = "131296"
  xl.Workbooks.Open "C:\ADMIN\OFFICE\Facturacion.xls", 0, True, 4, WPAS, WPAS
Return

Exit Sub
CANCELA:
  FRMIMP.Pantalla.Enabled = True
  FRMIMP.Pantalla.Caption = "Por &Pantalla"
  FRMIMP.lblproceso.Visible = False
  LOC_CANCELA = 0
  Set xl = Nothing
  Screen.MousePointer = 0

Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FRMIMP
 
End Sub

