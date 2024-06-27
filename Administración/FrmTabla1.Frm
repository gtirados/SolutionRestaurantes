VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmTabla1 
   AutoRedraw      =   -1  'True
   Caption         =   "Tablas "
   ClientHeight    =   6255
   ClientLeft      =   360
   ClientTop       =   1050
   ClientWidth     =   8760
   ControlBox      =   0   'False
   Icon            =   "FrmTabla1.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6255
   ScaleWidth      =   8760
   WindowState     =   2  'Maximized
   Begin VB.ListBox tralista 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      Left            =   3000
      Style           =   1  'Checkbox
      TabIndex        =   26
      Top             =   900
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox txtcia 
      Enabled         =   0   'False
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
      Left            =   8640
      TabIndex        =   24
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox subtabla 
      Enabled         =   0   'False
      Height          =   315
      Left            =   7440
      TabIndex        =   21
      Top             =   720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame FrmOtp 
      Caption         =   "Opciones :"
      Enabled         =   0   'False
      Height          =   2775
      Left            =   7440
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
      Begin VB.OptionButton OptT 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   2040
         Width           =   1335
      End
      Begin VB.OptionButton OptT 
         Caption         =   "&Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1335
      End
      Begin VB.OptionButton OptT 
         Caption         =   "E&ditar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox TXTSELE 
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
      Left            =   3960
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   7440
      TabIndex        =   13
      Top             =   4200
      Visible         =   0   'False
      Width           =   1815
      Begin VB.CommandButton cmdpoder 
         Caption         =   "Grid Tabla {UPDATE}"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdDBGRID 
         Caption         =   "&Execute SQL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.ComboBox LIST_TABLAS 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Ce&rrar"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   4800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   3600
      Top             =   5760
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos :"
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton cmdopciones 
         Caption         =   "&Opciones"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdtra 
         Caption         =   "Activar &Transacciones"
         Height          =   375
         Left            =   5040
         TabIndex        =   25
         Top             =   4200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin ComctlLib.ProgressBar ProgBar 
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   4680
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid GridT 
         Height          =   3855
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   6800
         _Version        =   393216
         FocusRect       =   2
         HighLight       =   0
         AllowUserResizing=   1
      End
      Begin VB.CommandButton CmdCodigo 
         Caption         =   "&Mostrar"
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   4200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Carga 
         Height          =   255
         Left            =   1200
         TabIndex        =   16
         Top             =   4680
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label LblRegistros 
         AutoSize        =   -1  'True
         Caption         =   "       "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   4440
         Width           =   795
      End
      Begin VB.Label LblCampos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   4680
         Width           =   1575
      End
   End
   Begin RichTextLib.RichTextBox Do_Tab 
      Height          =   975
      Left            =   4440
      TabIndex        =   11
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"FrmTabla1.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog dialogo 
      Left            =   1560
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label TXT2 
      Caption         =   "Compañia :"
      Height          =   255
      Left            =   7440
      TabIndex        =   23
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label TXT1 
      Caption         =   "Cod.Tabla :"
      Height          =   255
      Left            =   7440
      TabIndex        =   22
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Seleccion :"
      Height          =   255
      Left            =   3000
      TabIndex        =   19
      Top             =   120
      Width           =   855
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2280
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTabla1.frx":04C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTabla1.frx":065E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTabla1.frx":0770
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTabla1.frx":086A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Mensaje 
      AutoSize        =   -1  'True
      Caption         =   "       "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8040
      TabIndex        =   9
      Top             =   0
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "Tabla :"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "FrmTabla1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset
Dim PS_REP02 As rdoQuery
Dim llave_rep02 As rdoResultset
Dim bandera As Integer
Dim CORTA As String
Dim FIL_REG As Integer
Dim WPAS As Integer
Dim LOC_BAN As String * 1

Public Function tipo(Codigo As Integer) As String
Dim VAR As String
Select Case Codigo
Case 1
 VAR = " rdTypeCHAR"
Case 2
 VAR = "rdTypeNUMERIC"
Case 3
 VAR = "rdTypedecimal1"
Case 4
 VAR = "rdTypeINTEGER"
Case 5
 VAR = "rdTypeSMALLINT"
Case 6
 VAR = "rdTypeFLOAT"
Case 7
 VAR = "rdTypeREAL"
Case 8
 VAR = "rdTypeDOUBLE"
Case 9
 VAR = "rdTypeDATE"
Case 10
 VAR = "rdTypeTIME"
Case 11
 VAR = "rdTypeTIMESTAMP"
Case 12
 VAR = "rdTypeVARCHAR"
Case -1
 VAR = "rdTypeLONGVARCHAR"
Case -2
 VAR = "rdTypeBINARY"
Case -3
 VAR = "rdTypeVARBINARY"
Case -4
 VAR = "rdTypeLONGVARBINARY"
Case -5
 VAR = "rdTypeBIGINT"
Case -6
 VAR = "rdTypeTINYINT"
Case -7
 VAR = "rdTypeBIT"
End Select
tipo = VAR
End Function



Private Sub CmdCodigo_Click()
On Error GoTo SALE
Dim a, h, wBAN, R1 As Integer
Dim tempo As Integer
If Left(LIST_TABLAS.text, 1) = "<" Or Trim(LIST_TABLAS.text) = "" Then
  LIST_TABLAS.SetFocus
  Exit Sub
End If
If Left(CmdCodigo.Caption, 2) = "&D" Then
   CmdCodigo.Caption = "&Mostrar"
   WStop = True
   Exit Sub
Else
   CmdCodigo.Caption = "&Detener"
End If
Screen.MousePointer = 0
CmdCodigo.Enabled = True
CmdCodigo.SetFocus
FrmTabla1.GridT.Visible = False
FrmTabla1.GridT.Enabled = False
FrmOtp.Enabled = False
CmdSalir.Enabled = False
CmdDBGRID.Enabled = False
cmdtra.Enabled = False
LIST_TABLAS.Enabled = False
bandera = 0
If Trim(TXTSELE.text) = "" Then
  LOC_BAN = "N"
  If subtabla.text <> "" Or txtcia.text <> "" Then
   LOC_BAN = "S"
  End If
Else
  LOC_BAN = "S"
End If
CONECT_TAB LIST_TABLAS.text
If CORTA = "ERROR" Then
  Screen.MousePointer = 0
  CmdCodigo.Caption = "&Mostrar"
  LIST_TABLAS.Enabled = True
  TXTSELE.SetFocus
  cmdtra.Enabled = True
  Exit Sub
End If
Gen_llave.Requery
CAMPOS1 = Gen_llave.rdoColumns.Count
WStop = False
If Gen_llave.EOF Then
   If LOC_BAN = "S" Then
        MsgBox "No hay Datos en la Consulta .. Intente nuevamente . . ", 48, Pub_Titulo
        CmdCodigo.Caption = "&Mostrar"
        CmdSalir.Enabled = True
        cmdtra.Enabled = True
        LIST_TABLAS.Enabled = True
        GridT.Visible = False
        If TXTSELE.text = "" And txtcia.Enabled Then
          subtabla.SetFocus
          Screen.MousePointer = 0
          Exit Sub
        End If
        TXTSELE.SetFocus
        Screen.MousePointer = 0
        Exit Sub
   Else
        cmdtra.Enabled = True
        pub_mensaje = "NO Hay Registros... ¿Desea Agregar Registros... ?"
        Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
        If Pub_Respuesta = vbYes Then   ' El usuario eligió
          Screen.MousePointer = 0
          CmdCodigo.Caption = "&Mostrar"
          CmdSalir.Enabled = True
          Posi_Reg = 1
          OptT_KeyPress 1, 13
          Exit Sub
        End If
    End If
   CmdCodigo.Caption = "&Mostrar"
   CmdSalir.Enabled = True
   cmdtra.Enabled = True
   LIST_TABLAS.Enabled = True
   LIST_TABLAS.SetFocus
   GridT.Visible = False
   Screen.MousePointer = 0
  Exit Sub

End If
If Trim(LK_CODUSU) <> "ADMIN" And Trim(LIST_TABLAS.text) = "USUARIOS" Then
 CAMPOS1 = 2
Else
 CAMPOS1 = Gen_llave.rdoColumns.Count
End If
FIL_REG = Gen_llave.RowCount
FrmTabla1.GridT.Clear
GridT.Cols = CAMPOS1 + 1
GridT.Rows = FIL_REG + 1
GridT.RowHeight(0) = 400
GridT.ColWidth(0) = 1
FrmTabla1.ProgBar.Visible = True
tempo = CAMPOS1 + 1
GridT.Cols = tempo
For a = 1 To CAMPOS1
  If Left(Gen_llave.rdoColumns(a - 1).Name, 1) = "X" Then
    GridT.ColWidth(a) = 100
    GridT.TextMatrix(0, a) = "-"
  Else
    GridT.ColWidth(a) = 1500
    GridT.TextMatrix(0, a) = Gen_llave.rdoColumns(a - 1).Name '& "Nº : " & a
   GridT.Visible = False
  End If
Next a
LLENA_TABLA
WStop = False
CAMPOS1 = tempo - 1
GridT.Visible = True
LIST_TABLAS.Enabled = True
FrmTabla1.GridT.Enabled = True
cmdtra.Enabled = True
GridT.SetFocus
GridT.Col = 1
GridT.Row = 1
bandera = 1
If Trim(LK_CODUSU) <> "ADMIN" And Trim(LIST_TABLAS.text) = "USUARIOS" Then
 GridT.ColWidth(2) = 3000
End If
If Trim(LK_CODUSU) = "ADMIN" And Trim(LIST_TABLAS.text) = "USUARIOS" Then
 GridT.ColWidth(3) = 0
End If
LblRegistros.Caption = "Registro : " & FIL_REG
LblCampos.Caption = "Campo    : " & tempo - 1
CmdCodigo.Caption = "&Mostrar"
FrmOtp.Enabled = True
CmdSalir.Enabled = True
CmdDBGRID.Enabled = True
Screen.MousePointer = 0
CmdCodigo.Enabled = True
FrmOtp.Enabled = True
FrmOtp.Enabled = True
If Trim(FrmTabla1.LIST_TABLAS.text) = "PARGEN" Then
    For a = 0 To 20
      MDIForm1.menuCia(a).Visible = True
      MDIForm1.menuCia(a).Checked = False
    Next a
    PS_PAR(0) = " "
    par.Requery
    par.MoveFirst
    h = 0
    Do Until par.EOF
        WCOD = par!PAR_CODCIA
        If LK_CODCIA = par!PAR_CODCIA Then wBAN = h
        MDIForm1.menuCia(h).Caption = WCOD & "-" & Trim(par!PAR_NOMBRE)
        h = h + 1
        par.MoveNext
    Loop
    For a = h To 20
      MDIForm1.menuCia(a).Visible = False
    Next a
    MDIForm1.menuCia(wBAN).Checked = True
End If
 
Exit Sub
SALE:
If Err.Number = 30006 Then
   MsgBox "No se pudo Mostrar Datos, Filtre para obtener menos datos ", 48, Pub_Titulo
Else
   MsgBox Err.Description
   Resume Next
End If
TXTSELE.SetFocus
CmdCodigo.Enabled = True
CmdCodigo.Caption = "&Mostrar"
CmdSalir.Enabled = True
LIST_TABLAS.Enabled = True

End Sub

Private Sub CmdCodigo_KeyPress(KeyAscii As Integer)
LIST_TABLAS.Enabled = True
LIST_TABLAS.SetFocus
SendKeys "%{UP}"
End Sub

Private Sub CmdDBGRID_Click()
On Error GoTo SALE
Dim wvalor
Dim valor As Integer
If LK_CODUSU <> "ADMIN" Then
  MsgBox "Solo el Administrador Puede Ingresar ...Consultar ! ", 48, Pub_Titulo
  Exit Sub
End If
    pub_mensaje = "Solo para conocedores, expertos, ingenieros  en SQL SERVER 6.5 pueden Utilizar esta opción, mejor Consulte a su Administrador de Sistemas. Cualquier perdida  de DATOS sera imposible de recuperar (Realizar su RESTARE)."
    Pub_Titulo = "COMMAND EN SQL SERVER 6.5 "
    valorpred = "UPDATE {tabla} SET {asignas}  WHERE  {condición}"
    wvalor = InputBox(pub_mensaje, Pub_Titulo, valorpred)
    If wvalor = "" Then
      Exit Sub
    End If
    Screen.MousePointer = 11
    CN.Execute wvalor, rdExecDirect
    Screen.MousePointer = 0
    MsgBox "Su Linea de Comando fué Ejecutada . ", 48, Pub_Titulo
    GridT.Visible = False
    CmdCodigo.Enabled = True
    CmdCodigo.SetFocus
Exit Sub
SALE:
   Screen.MousePointer = 0
   MsgBox Err.Description, vbCritical, Pub_Titulo
    
End Sub

Private Sub CmdDBGRID_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    LIST_TABLAS.SetFocus
End If
End Sub

Private Sub cmdopciones_Click()
On Error GoTo SALE
If LK_CODUSU = "ADMIN" Then
 Load frmopciones
 frmopciones.Show 1
End If
Exit Sub
SALE:
MsgBox "Falta el Formulario ..."
End Sub

Private Sub cmdpoder_Click()

If LK_CODUSU <> "ADMIN" Then
  MsgBox "Solo el Administrador Puede Ingresar ...Consultar ! ", 48, Pub_Titulo
  Exit Sub
End If
If Trim(LIST_TABLAS.text) = "" Then
 MsgBox "Seleccionar Tabla ..", 48, Pub_Titulo
 Exit Sub
End If
wTABLA = Trim(LIST_TABLAS.text)
Load Form1
Form1.Show 1
End Sub

Private Sub CmdSalir_Click()
On Error GoTo SALE
 Gen_llave.Cancel
 Gen_llave.Close
 Unload FrmTabla1
 Exit Sub
SALE:
 Unload FrmTabla1
End Sub

Private Sub cmdtra_Click()
Dim WCOD As String * 5
Screen.MousePointer = 11
pub_cadena = "SELECT TRA_KEY, TRA_DESCRIPCION, TRA_FLAG_ACTIVO FROM TRANSACCION ORDER BY TRA_KEY,TRA_DESCRIPCION"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
llave_rep01.Requery
If llave_rep01.EOF Then
 Screen.MousePointer = 0
 MsgBox "NO Existe Transacciones Consulte a su Proveedor !!!!", 48, Pub_Titulo
 Exit Sub
End If
ProgBar.Max = llave_rep01.RowCount
DoEvents
FrmTabla1.Carga.Caption = "Generando Elementos ..."
ProgBar.Visible = True
ProgBar.Value = 0
tralista.Clear
fila = -1
Do Until llave_rep01.EOF
  fila = fila + 1
  ProgBar.Value = ProgBar.Value + 1
  WCOD = llave_rep01!TRA_KEY
  tralista.AddItem WCOD & " " & Trim(llave_rep01!tra_descripcion)
  If llave_rep01!TRA_FLAG_ACTIVO = "A" Then
   tralista.Selected(fila) = True
  Else
   tralista.Selected(fila) = False
  End If
  llave_rep01.MoveNext
Loop
Screen.MousePointer = 0
FrmTabla1.Carga.Caption = ""
tralista.Visible = True
tralista.ListIndex = 0
ProgBar.Visible = False
tralista.SetFocus

End Sub

Private Sub Form_Activate()
If FRM_STATUS = "1" Then
    ACEPTA = ""
    OP_FORM = "X"
    FRM_STATUS = "0"
    Unload FrmTabla2
    Unload frmUSUARIOS
    Unload FRMTRANS
End If
OptT(0).Value = False
OptT(1).Value = False
OptT(2).Value = False
If ACEPTA = "F" Then
 ' CmdCodigo_Click   'Para mostrar lo editado o agregado
 If OP_FORM = "D" Then
  'msgbox "Registro Actualizado..  " & espacio & "Si desea ver lo Modificado Seleccione Mostrar ", 48, "Grupo Empresarial"
  FrmTabla1.GridT.Visible = True
  FrmTabla1.GridT.Col = 1
  FrmTabla1.GridT.SetFocus
  ACEPTA = ""
  OP_FORM = "X"
  Exit Sub
 ElseIf OP_FORM = "A" Then
  ' msgbox "Registro Agregado(s).. :  " & espacio & " Si desea ver lo Modificado Seleccione Mostrar ", 48, pub_titulo
   FrmTabla1.GridT.Col = 1
   If FrmTabla1.GridT.Visible Then
      FrmTabla1.GridT.SetFocus
   End If
   ACEPTA = ""
   OP_FORM = "X"
   Exit Sub
 End If
 CmdCodigo.Enabled = True
 Timer1.Enabled = True
 FrmTabla1.GridT.Row = Posi_Reg
 ACEPTA = ""
 OP_FORM = "X"
Exit Sub
End If
If OP_FORM = "D" Or OP_FORM = "A" Then
  'msgbox "Proceso  Cancelado ...", 48, pub_titulo
  ACEPTA = ""
  OP_FORM = "X"
  FrmTabla1.GridT.Col = 1
  FrmTabla1.GridT.SetFocus
  Exit Sub
  'verificar
  If FrmTabla1.GridT.Visible And FrmTabla1.GridT.Enabled Then
      FrmTabla1.GridT.SetFocus
  Else
      FrmTabla1.LIST_TABLAS.SetFocus
  End If
  Timer1.Enabled = True
  ACEPTA = ""
  OP_FORM = "X"
  Exit Sub
End If
ACEPTA = ""
OP_FORM = "X"

End Sub

Private Sub Form_Load()
WStop = False
bandera = 0
WPAS = 0
DoEvents
FrmTabla1.Show
Dim t As Integer
CORTA = "PASA"
OP_FORM = "X"
Dim cad As String
If Trim(LK_CODUSU) <> "ADMIN" Then
   cmdpoder.Visible = False
   CmdDBGRID.Visible = False
   LIST_TABLAS.AddItem "USUARIOS"
Else
   If CN.rdoTables.Count = 0 Then
      MsgBox " Esta bloqueado...", 48, Pub_Titulo
      LIST_TABLAS.SetFocus
      Exit Sub
   End If
    For t = 0 To CN.rdoTables.Count - 1
        DoEvents
        LIST_TABLAS.AddItem CN.rdoTables(t).Name
    Next
   cmdopciones.Visible = True
End If
'LIST_TABLAS.text = LIST_TABLAS.List(0)
FrmOtp.Visible = True
Frame1.Visible = True
'Frame2.Visible = True
Frame3.Visible = True
CmdCodigo.Visible = True
CmdSalir.Visible = True
TXTSELE.Visible = True
LIST_TABLAS.Visible = True
txtcia.Visible = True
subtabla.Visible = True
LIST_TABLAS.SetFocus
SendKeys "%{UP}"
End Sub



Public Sub LLENA_TABLA()
Dim FI, a As Integer
On Error GoTo Mens_Err
Gen_llave.MoveFirst
FI = 0

ProgBar.Max = Gen_llave.RowCount
DoEvents
FrmTabla1.Carga.Caption = "Generando Elementos ..."
FrmTabla1.Carga.Visible = True
Do Until Gen_llave.EOF
  If WStop Then
     Exit Do
  End If
  FI = FI + 1
  ProgBar.Value = FI
  For a = 1 To CAMPOS1
   If Not Gen_llave(a - 1) = "" Then
      DoEvents
      FrmTabla1.GridT.TextMatrix(FI, a) = CStr(Gen_llave(a - 1))
   End If
   If IsNumeric(Gen_llave(a - 1)) Then
      FrmTabla1.GridT.ColAlignment(a) = 1
   End If
  Next a
  Gen_llave.MoveNext
Loop
ProgBar.Value = 0
FrmTabla1.Carga.Visible = False
FrmTabla1.ProgBar.Visible = False
Mens_Err:
If Err.Number = 40036 Then
   MsgBox "NO se puede Abrir la Tabla.... Campos sin enlazar ", 48, Pub_Titulo
   End
End If
End Sub



Private Sub GridT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If GridT.text = "" Then
    Exit Sub
 End If
 Posi_Reg = GridT.Row
 Gen_llave.AbsolutePosition = Posi_Reg
 FrmTabla1.GridT.Col = 0
 FrmTabla1.GridT.Row = FrmTabla1.GridT.Row
 FrmTabla1.GridT.ColSel = FrmTabla1.GridT.Cols - 1
 FrmTabla1.GridT.RowSel = FrmTabla1.GridT.Row

 
' OptT(0).SetFocus
 If OptT(0).Enabled Then
  OptT(0).SetFocus
 End If
End If
If KeyAscii = 27 Then
 LIST_TABLAS.SetFocus
 SendKeys "%{UP}"
 FrmTabla1.GridT.Visible = False
End If
Mensaje.Caption = " "
End Sub

Private Sub GridT_KeyUp(KeyCode As Integer, Shift As Integer)
Dim a, queda As Integer
Dim t1, t2 As Integer
Dim W1, w2 As Integer
Dim wblo
wblo = ""
queda = GridT.Col
If KeyCode = 113 Or KeyCode = 114 Then
   t1 = 0
   DoEvents
   LIST_TABLAS.Enabled = False
   DoEvents
   CmdCodigo.Enabled = False
   CmdSalir.Enabled = False
   FrmTabla1.GridT.Visible = False
   DoEvents
   FrmTabla1.Carga.Visible = True
   DoEvents
   FrmTabla1.Carga.Caption = "Un Momento ..."
   DoEvents
   Do Until t1 = FrmTabla1.GridT.Rows - 1
      t1 = t1 + 1
      bandera = 0
      FrmTabla1.GridT.Row = t1
      GridT.Col = 1
      W1 = Len(GridT.text)
      wblo = Trim(GridT.text)
      GridT.Col = 2
      w2 = Len(GridT.text)
      If W1 = 0 And w2 = 0 Then
        Exit Do
      End If
      If Trim(FrmTabla1.LIST_TABLAS.text) = "USUARIOS" And wblo = "ADMIN" Or wblo = "SUPERVISOR" Then
        GoTo sigue
      End If
      
      For a = 0 To FrmTabla1.GridT.Cols - 1
         FrmTabla1.GridT.Col = a
         If KeyCode = 113 Then
            FrmTabla1.GridT.CellBackColor = vbScrollBars
         ElseIf KeyCode = 114 Then
            FrmTabla1.GridT.CellBackColor = vbWhite
         End If
      Next a
sigue:
   Loop
   FrmTabla1.Carga.Visible = False
   If KeyCode = 113 Then
      bandera = 9
   ElseIf KeyCode = 114 Then
      bandera = 0
   End If
  
   LIST_TABLAS.Enabled = True
   CmdCodigo.Enabled = True
   CmdSalir.Enabled = True
   FrmTabla1.GridT.Visible = True
   FrmTabla1.GridT.Col = queda
   FrmTabla1.GridT.Row = 2
   FrmTabla1.GridT.SetFocus
End If
If KeyCode = 32 Then
   If Trim(LIST_TABLAS.text) = "USUARIOS" And Trim(GridT.TextMatrix(GridT.Row, 1)) = "ADMIN" Or Trim(GridT.TextMatrix(GridT.Row, 1)) = "SUPERVISOR" Then
      MsgBox "NO Procese , por ser Administrador..", 48, Pub_Titulo
       GridT.SetFocus
      Exit Sub
   End If
   GridT.Col = 1
   t1 = Len(GridT.text)
   GridT.Col = 2
   t2 = Len(GridT.text)
   If t1 = 0 Or t2 = 0 Then
      Exit Sub
   End If
   
   If FrmTabla1.GridT.CellBackColor = vbScrollBars Then
      For a = 0 To FrmTabla1.GridT.Cols - 1
        FrmTabla1.GridT.Col = a
        FrmTabla1.GridT.CellBackColor = vbWhite
      Next a
    Else
      For a = 0 To FrmTabla1.GridT.Cols - 1
        FrmTabla1.GridT.Col = a
        FrmTabla1.GridT.CellBackColor = vbScrollBars
      Next a
      bandera = 9
    End If
    FrmTabla1.GridT.Col = queda
    FrmTabla1.GridT.SetFocus
End If
If KeyCode = 46 And bandera = 9 Then
  pub_mensaje = " ¿Desea Eliminar lo Seleccionado ... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbYes Then   ' El usuario eligió
   GridT.Visible = False
   Screen.MousePointer = 11
   WASIGNA
   Screen.MousePointer = 0
  Else
  FrmTabla1.GridT.Col = queda
  FrmTabla1.GridT.SetFocus
  End If
End If
End Sub

Private Sub LIST_TABLAS_GotFocus()
    FrmTabla1.GridT.Visible = False
    CmdCodigo.Enabled = True
    FrmTabla1.TXTSELE.text = ""
End Sub

Private Sub LIST_TABLAS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdtra.Visible = False
  If Trim(LIST_TABLAS.text) = "TABLAS" Then
    txtcia.Enabled = True
    subtabla.Enabled = True
  ElseIf Trim(LIST_TABLAS.text) = "TRANSACCION" Then
    cmdtra.Visible = True
  Else
    txtcia.Enabled = False
    subtabla.Enabled = False
  End If
  If LIST_TABLAS.text <> "" Then
     CmdCodigo.Enabled = True
     TXTSELE.SetFocus
  End If
 End If
End Sub


Private Sub OptT_DblClick(Index As Integer)
GridT_KeyPress 13
OptT_KeyPress Index, 13
End Sub

Private Sub OptT_GotFocus(Index As Integer)
If GridT.Visible Then
 FrmOtp.Enabled = True
Else
 FrmOtp.Enabled = False
End If


End Sub



Private Sub OptT_KeyPress(Op As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then
 FrmTabla1.GridT.Col = 1
 FrmTabla1.GridT.SetFocus
 Exit Sub
End If
On Error GoTo GG
If KeyAscii = 13 Then
  If Op = 0 And FrmTabla1.GridT.Col >= 0 Then
     OP_FORM = "D"
     GoTo Carga
     Exit Sub
  ElseIf Op = 1 Then
     OP_FORM = "A"
     GoTo Carga
     Exit Sub
  ElseIf Op = 2 And FrmTabla1.GridT.Col >= 0 Then
     GoTo Eli
     Exit Sub
  End If
  If OptT(1).Value Then
     OP_FORM = "A"
     GoTo Carga
     Exit Sub
  End If
  If OptT(1).Value And FrmTabla1.GridT.Col >= 0 Then
   OP_FORM = "D"
   GoTo Carga
   Exit Sub
  End If
Exit Sub
Eli:
  If OptT(2).Value Or Op = 2 Then
    GridT.Col = 1
    If Trim(LIST_TABLAS.text) = "USUARIOS" And Trim(GridT.text) = "ADMIN" Or Trim(GridT.text) = "SUPERVISOR" Then
        MsgBox "Imposible ELIMINAR al Administrador ....", 48, Pub_Titulo
        GridT.Col = 0
        Exit Sub
    End If
    OP_FORM = "E"
    pub_mensaje = " ¿Desea Eliminar el Registro... ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbYes Then   ' El usuario eligió
       Gen_llave.AbsolutePosition = Posi_Reg
       Gen_llave.Delete
       Gen_llave.MoveNext
       CmdCodigo_Click
       Exit Sub
    End If
   Exit Sub
   End If
End If
Carga:
  Cta_Add = 0
  Timer1.Enabled = False
  ACEPTA = "V"
  wTABLA = LIST_TABLAS.text
  Screen.MousePointer = 11
  FRM_STATUS = "1"
  Select Case LIST_TABLAS.text
   Case "TRANSACCION"
'     DoEvents
     FrmTabla1.Hide
     FRMTRANS.Show
     DoEvents
     Load FRMTRANS
   Case "USUARIOS"
     FrmTabla1.Hide
     frmUSUARIOS.Show
     Load frmUSUARIOS
     
   Case Else
     FrmTabla1.Hide
     Load FrmTabla2
  End Select
Screen.MousePointer = 0
Exit Sub
GG:
  MsgBox Err.Description & " " & "  "
  Screen.MousePointer = 0
  FRM_STATUS = "0"
  ACEPTA = ""
  OP_FORM = "X"
  
End Sub



Public Sub CONECT_TAB(tabla As String)
Dim wt_cia, wt_tipreg
'On Error GoTo ERRO
If Not WPAS = 0 Then
    WPAS = 1
    Gen_llave.Cancel
  '  Gen_llave.Close
End If
pub_cadena = "SELECT * FROM "
On Error GoTo ERRO
If tabla = "TABLAS" Then
  If Len(Trim(FrmTabla1.TXTSELE)) = 0 Then
     wt_cia = txtcia.text
     wt_tipreg = subtabla.text
     If wt_cia <> "" And wt_tipreg <> "" Then
         pub_cadena = pub_cadena + tabla & " WHERE  TAB_CODCIA = '" & wt_cia & "' AND TAB_TIPREG = " & wt_tipreg
     End If
     If wt_cia <> "" And wt_tipreg = "" Then
         pub_cadena = pub_cadena + tabla & " WHERE  TAB_CODCIA = '" & wt_cia & "' "
     End If
     If wt_cia = "" And wt_tipreg <> "" Then
         pub_cadena = pub_cadena + tabla & " WHERE  TAB_TIPREG = " & wt_tipreg
     End If
     If wt_cia = "" And wt_tipreg = "" Then
     Else
      GoTo DALET
     End If
  End If
End If
If Len(Trim(FrmTabla1.TXTSELE)) = 0 Then
  pub_cadena = pub_cadena + tabla
  If LK_CODUSU <> "ADMIN" And Trim(LIST_TABLAS.text) = "USUARIOS" Then
   pub_cadena = pub_cadena + " WHERE  USU_KEY <> 'ADMIN'"
  End If
Else
  pub_cadena = pub_cadena + tabla & " WHERE " & Trim(FrmTabla1.TXTSELE)
  If LK_CODUSU <> "ADMIN" And Trim(LIST_TABLAS.text) = "USUARIOS" Then
   pub_cadena = pub_cadena + " AND USU_KEY <> 'ADMIN'"
  End If
End If
DALET:
Set LLAVE = CN.CreateQuery("", pub_cadena)
DoEvents
Set Gen_llave = LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
CORTA = "PASA"
Exit Sub
ERRO:
MsgBox Err.Description, 48, Pub_Titulo
CmdSalir.Enabled = True
 CORTA = "ERROR"
'Dim er As rdoError
    'Debug.Print Err, Error
    'For Each er In rdoErrors
    '    Debug.Print er.Description, er.Number
    'Next er
 
 'If CORTA <> "ERROR" Then
   'msgbox rdErr.Description '"NO EXISTE  TABLA ..!"

 'End If
End Sub

Public Sub Impri()
Dim i
Dim MISIZE As Currency
Static t
Dim WT As String
Dim text As String
Dim a As Integer
Dim cuenta As Integer
Dim suma As Integer
a = 0
t = 2
suma = 0
cuenta = cuenta + 1
Printer.CurrentX = 3500
WT = "TABLA : " & Mid(Gen_llave.Name, 15)
Printer.Print WT
Printer.Print " "
Printer.Print " "
Printer.FontSize = 10
Printer.FontName = "Arial"
For i = 0 To Gen_llave.rdoColumns.Count - 1
 a = a + 1
 Printer.CurrentX = 200
 Printer.Print Gen_llave.rdoColumns(i).Name;
 Printer.CurrentX = 3000
 Printer.Print tipo(Gen_llave.rdoColumns(i).Type);
 Printer.CurrentX = 5000
 MISIZE = Gen_llave.rdoColumns(i).Size
 Printer.Print MISIZE
Next i
Printer.Print "Total :."; a
Printer.EndDoc

End Sub

Public Sub WASIGNA()
Dim Q As Integer
Dim R1, R3 As Integer
Dim R2, wavisa As Integer
wavisa = 0
Q = 0
R1 = -1
R2 = -1
GridT.Row = 1
DoEvents
FrmTabla1.Carga.Visible = True
DoEvents
FrmTabla1.Carga.Caption = "Eliminando ..."
DoEvents
FrmTabla1.ProgBar.Min = 0
FrmTabla1.ProgBar.Max = GridT.Rows - 1
FrmTabla1.ProgBar.Value = 0
FrmTabla1.ProgBar.Visible = True

Do Until (R1 = 0 And R2 = 0) Or GridT.Row = GridT.Rows - 1
   Q = Q + 1
   GridT.Row = Q
   GridT.Col = 1
   R1 = Len(GridT.text)
   GridT.Col = 2
   R2 = Len(GridT.text)
   bandera = 0
   If GridT.CellBackColor = vbScrollBars Then
      Gen_llave.AbsolutePosition = Q
      Gen_llave.Delete
      wavisa = 1
  End If
  FrmTabla1.ProgBar.Value = Q
Loop
FrmTabla1.ProgBar.Value = GridT.Rows - 1
FrmTabla1.Carga.Visible = False
FrmTabla1.ProgBar.Visible = False
If wavisa = 1 Then
 GridT.Visible = False
 LIST_TABLAS.Enabled = True
 CmdCodigo.Enabled = True
 CmdCodigo.Caption = "&Mostrar"
 CmdCodigo.SetFocus
Else
  GridT.Row = 1
  GridT.Col = 1
End If
bandera = 0
End Sub

Private Sub subtabla_GotFocus()
TXTSELE.text = ""
End Sub

Private Sub subtabla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 txtcia.SetFocus
End If
End Sub

Private Sub tralista_Click()
'If Not ProgBar.Visible Then
' tralista_KeyDown 32, 0
'End If
End Sub

Private Sub tralista_ItemCheck(Item As Integer)
If ProgBar.Visible Then
 Exit Sub
End If
 If Val(Left(tralista.text, 5)) <= 0 Then
   Exit Sub
 End If
 SQ_OPER = 1
 PUB_CODTRA = Val(Left(tralista.text, 5))
 LEER_TRA_LLAVE
 If tra_llave.EOF Then
   MsgBox "Error de Llave de Transacción ... Verificar .. ", 48, Pub_Titulo
   Exit Sub
 End If
 Screen.MousePointer = 11
 tra_llave.Edit
 If tralista.Selected(tralista.ListIndex) Then
   tra_llave!TRA_FLAG_ACTIVO = "A"
 Else
   tra_llave!TRA_FLAG_ACTIVO = " "
 End If
 tra_llave.Update
 Screen.MousePointer = 0
End Sub

Private Sub tralista_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  tralista.Visible = False
  TXTSELE.SetFocus
  Exit Sub
End If
End Sub

Private Sub tralista_LostFocus()
tralista.Visible = False
End Sub

Private Sub txtcia_GotFocus()
TXTSELE.text = ""
End Sub

Private Sub txtcia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 CmdCodigo_Click
End If
End Sub

Private Sub TXTSELE_GotFocus()
subtabla.text = ""
txtcia.text = ""
End Sub

Private Sub TXTSELE_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  LIST_TABLAS.SetFocus
  Exit Sub
End If
If KeyAscii = 13 Then
  If TXTSELE.text <> "" Then
    CmdCodigo_Click
    Exit Sub
  ElseIf subtabla.Enabled Then
    subtabla.SetFocus
    Exit Sub
  End If
  CmdCodigo_Click
End If
End Sub
