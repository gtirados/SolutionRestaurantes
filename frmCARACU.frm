VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCaracu 
   Caption         =   "Letras de Clientes ..."
   ClientHeight    =   6060
   ClientLeft      =   285
   ClientTop       =   1320
   ClientWidth     =   9480
   Icon            =   "frmCARACU.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   6060
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   2055
      Begin VB.OptionButton Opt 
         Caption         =   "&Activas"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton Opt 
         Caption         =   "&Todas"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
      Begin VB.ComboBox CmbCP 
         Height          =   315
         ItemData        =   "frmCARACU.frx":0442
         Left            =   120
         List            =   "frmCARACU.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Opc.de  Letras"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Seleccionar"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame F4 
      Caption         =   "Historia :"
      Height          =   3975
      Left            =   3720
      TabIndex        =   5
      Top             =   1680
      Width           =   5655
      Begin MSFlexGridLib.MSFlexGrid GridHistoria 
         Height          =   3615
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   6376
         _Version        =   327680
         GridLineWidth   =   2
      End
   End
   Begin VB.Frame F1 
      Height          =   1575
      Left            =   2160
      TabIndex        =   3
      Top             =   0
      Width           =   7215
      Begin VB.CommandButton CmdEstadistica 
         Caption         =   "&Estadistica Grafica"
         Height          =   375
         Left            =   2400
         TabIndex        =   14
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton CmdRetornar 
         Caption         =   "&Retornar"
         Height          =   375
         Left            =   4440
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton CmdMostrar 
         Caption         =   "&Mostrar"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox TxtCodigo 
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
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label LblNombreCli 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo "
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame F3 
      Caption         =   "Letras :"
      Height          =   3975
      Left            =   0
      TabIndex        =   4
      Top             =   1680
      Width           =   3615
      Begin MSFlexGridLib.MSFlexGrid GridLetras 
         Height          =   3615
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   6376
         _Version        =   327680
         GridLineWidth   =   2
      End
   End
End
Attribute VB_Name = "FrmCaracu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmbCP_Click()
FrmCaracu.GenClie1.CLI_CP = Left(CmbCP.text, 1)
End Sub

Private Sub CmbCP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtCodigo.SetFocus
End If
End Sub

Private Sub CmdMostrar_Click()
Dim vESTADO As Integer
SQ_OPER = 1
PUB_CP = Left(CmbCP.text, 1)
PUB_CODCLIE = Val(TxtCodigo.text)
PUB_CODCIA = LK_CODCIA
LEER_CLI_LLAVE
If cli_llave.EOF Then
    MsgBox "No Existe Cliente ..!!!", 48, WS_TITULO
    Azul TxtCodigo, TxtCodigo
    Exit Sub
End If
GridLetras.Clear
'ACTIVAS O TODAS..
If Opt(0).Value Then
    vESTADO = 1
Else
    vESTADO = 0
End If


LLENA_LETRAS TxtCodigo.text, Left(CmbCP.text, 1), vESTADO, GridLetras
GridLetras.Enabled = True
GridLetras.Col = 1
GridLetras.Row = 1
GridLetras.SetFocus
End Sub

Private Sub CmdRetornar_Click()
 Unload FrmCaracu
End Sub

Private Sub Form_Load()
Dim i As Integer
GenClie1.DSN = PUB_DSN
CmbCP.ListIndex = 0
Opt(1).Value = True
'PARA GRIDLETRAS
GridLetras.Rows = 10
GridLetras.Cols = 10
GridLetras.ColWidth(0) = 1000
GridLetras.ColWidth(1) = 300
GridLetras.ColWidth(2) = 1100
GridLetras.ColWidth(3) = 500
GridLetras.ColWidth(4) = 1200
GridLetras.ColWidth(5) = 800
GridLetras.ColWidth(6) = 1000
GridLetras.ColWidth(7) = 800
GridLetras.ColWidth(8) = 800
GridLetras.ColWidth(9) = 2000
'For I = 1 To GridLetras.Cols - 1
'    GridLetras.ColAlignment(I) = 0
'Next I
GridLetras.ColAlignment(2) = 1
'GridLetras.CellAlignment(1) = 4
CABE_LETRAS
'PARA GRIDHISTORIA
GridHistoria.Rows = 10
GridHistoria.Cols = 10
CABE_HISTORIA
GridHistoria.ColWidth(0) = 1000
GridHistoria.ColWidth(1) = 300
GridHistoria.ColWidth(2) = 1000
GridHistoria.ColWidth(3) = 1000
GridHistoria.ColWidth(4) = 1000
GridHistoria.ColWidth(5) = 1000
GridHistoria.ColWidth(6) = 1000
GridHistoria.ColWidth(7) = 1000
GridHistoria.ColWidth(8) = 2000
GridHistoria.ColAlignment(3) = 1
GridHistoria.ColAlignment(4) = 1
End Sub

Private Sub GenClie1_MyKeyGrid1(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    GenClie1.Visible = False
    Azul TxtCodigo, TxtCodigo
End If
If KeyCode = 32 Then
    If Len(GenClie1.Codigo) = 0 Then
        Exit Sub
    End If
    GenClie1.Visible = False
    TxtCodigo.text = GenClie1.Codigo
    LblNombreCli.Caption = GenClie1.Nombre
    CmdMostrar.SetFocus
End If

End Sub

Private Sub GenClie1_MyKeyLeido(KeyAscii As Integer)
If KeyAscii = 27 Then
    GenClie1.Visible = False
    Azul TxtCodigo, TxtCodigo
End If
End Sub

Private Sub GenClie1_MyKeyLeido2(KeyAscii As Integer)
If KeyAscii = 27 Then
    GenClie1.Visible = False
    Azul TxtCodigo, TxtCodigo
End If
End Sub

Private Sub GridHistoria_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  GridHistoria.Clear
  GridLetras.Col = 1
  GridHistoria.Enabled = False
  GridLetras.SetFocus
End If
End Sub

Private Sub GridHistoria_LostFocus()
  GridHistoria.Clear
  GridLetras.Col = 1
  GridHistoria.Enabled = False
  GridLetras.SetFocus
End Sub

Private Sub GridLetras_DblClick()
GridLetras_KeyPress 13
End Sub

Private Sub GridLetras_KeyPress(KeyAscii As Integer)
Dim CP As String
Dim CODCLIE As String
Dim TIPDOC As String
Dim FECHA_ING As String
Dim NUMDOC As String

If KeyAscii = 27 Then
   GridLetras.Clear
   GridLetras.Enabled = False
   Azul TxtCodigo, TxtCodigo
   Exit Sub
End If
If KeyAscii = 13 Then
    GridLetras.Col = 1
    If GridLetras.text = "" Then
       Exit Sub
    End If
     GridHistoria.Enabled = True
    'LLENA TODA LA HISTORIA
    CP = Left(CmbCP.text, 1)
    GridLetras.Col = 1
    CODCLIE = TxtCodigo.text
    TIPDOC = GridLetras.text
    GridLetras.Col = 4
    FECHA_ING = GridLetras.text
    GridLetras.Col = 5
    NUMDOC = GridLetras.text
    LLENA_HISTORIA CODCLIE, CP, TIPDOC, FECHA_ING, NUMDOC, GridHistoria
    GridLetras.Col = 1
    GridLetras.Row = GridLetras.Row
    GridLetras.ColSel = GridLetras.Cols - 1
    GridLetras.RowSel = GridLetras.Row
    GridHistoria.Col = 2
    GridHistoria.Row = 1
    GridHistoria.SetFocus
    ';MsgBox "falta"
   
End If
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TxtCodigo.text = "" Then
        GenClie1.Active = True
        GenClie1.procesa
        GenClie1.Visible = True
        GenClie1.SetFocus
    Else
        SQ_OPER = 1
        PUB_CP = Left(CmbCP.text, 1)
        PUB_CODCLIE = Val(TxtCodigo.text)
        PUB_CODCIA = LK_CODCIA
        LEER_CLI_LLAVE
        If cli_llave.EOF Then
            MsgBox "No Existe Cliente ..!!!", 48, WS_TITULO
            Azul TxtCodigo, TxtCodigo
            Exit Sub
        End If
        LblNombreCli.Caption = cli_llave!CLI_NOMBRE
        CmdMostrar.SetFocus
    End If

End If
End Sub

Public Sub LLENA_LETRAS(CODI As String, wCP As String, WESTADO As Integer, WGRID As MSFlexGrid)
Dim WCOLOR As Double
Dim CUENTA As Integer
Dim FLAG As Integer
Dim NUMERO As String * 3
SQ_OPER = 3
PUB_CP = wCP
PUB_CODCLIE = Val(CODI)
LEER_CAR_LLAVE
If car_menor.EOF Then
    WGRID.Col = 0
    WGRID.Row = 1
    WGRID.text = "No hay Letras.."
    Exit Sub
End If
'WESTADO  , 1 ACTIVAS Y 0 SON TODAS

WCOLOR = vbWhite '** COLOR DE ACTIVAS


CABE_LETRAS
CUENTA = 0
WGRID.Rows = 2
CUENTA = CUENTA + 1
WGRID.Row = CUENTA
WGRID.Col = 2
WGRID.text = "ACTIVAS        "
WGRID.Rows = WGRID.Rows + 1
Do Until car_menor.EOF
   If car_menor!car_situacion <> 9 Then
        CUENTA = CUENTA + 1
        WGRID.Row = CUENTA
        WGRID.Col = 0
        NUMERO = CUENTA
        WGRID.text = car_menor!CAR_fecha_vcto
        WGRID.Col = 1
        WGRID.CellBackColor = WCOLOR
        WGRID.text = Trim(car_menor!CAR_TIPDOC)
        WGRID.Col = 2
        WGRID.CellBackColor = WCOLOR
        WGRID.text = car_menor!CAR_IMPORTE
        WGRID.Col = 3
        WGRID.CellBackColor = WCOLOR
        WGRID.text = car_menor!car_situacion
        WGRID.Col = 4
        WGRID.CellBackColor = WCOLOR
        WGRID.text = car_menor!car_fecha_ingr
        WGRID.Col = 5
        WGRID.CellBackColor = WCOLOR
        WGRID.text = car_menor!car_numdoc
        WGRID.Col = 6
        WGRID.CellBackColor = WCOLOR
        WGRID.text = car_menor!car_int_8dias
        WGRID.Col = 7
        WGRID.CellBackColor = WCOLOR
        WGRID.text = Nulo_Valors(car_menor!car_numser)
        WGRID.Col = 8
        WGRID.CellBackColor = WCOLOR
        WGRID.text = Nulo_Valors(car_menor!car_numfac)
        WGRID.Col = 9
        WGRID.CellBackColor = WCOLOR
        WGRID.text = Nulo_Valors(car_menor!car_concepto)
        WGRID.Rows = WGRID.Rows + 1
    End If
    car_menor.MoveNext
Loop
If WESTADO = 1 Then
   Exit Sub
End If
WCOLOR = vb3DLight    '** COLOR DE CANCELADAS
WGRID.Rows = WGRID.Rows + 1
CUENTA = CUENTA + 1
WGRID.Row = CUENTA
WGRID.Col = 2
WGRID.text = "CANCELADAS"
car_menor.MoveFirst
Do Until car_menor.EOF
   If car_menor!car_situacion = 9 Then
        CUENTA = CUENTA + 1
        WGRID.Row = CUENTA
        WGRID.CellBackColor = WCOLOR
        WGRID.Col = 0
        NUMERO = CUENTA
        WGRID.text = car_menor!CAR_fecha_vcto
        WGRID.Col = 1
        WGRID.CellBackColor = WCOLOR
        WGRID.text = Trim(car_menor!CAR_TIPDOC)
        WGRID.Col = 2
        WGRID.CellBackColor = WCOLOR
        WGRID.text = car_menor!CAR_IMPORTE
        WGRID.Col = 3
        WGRID.CellBackColor = WCOLOR
        WGRID.text = car_menor!car_situacion
        WGRID.Col = 4
        WGRID.CellBackColor = WCOLOR
        WGRID.text = car_menor!car_fecha_ingr
        WGRID.Col = 5
        WGRID.CellBackColor = WCOLOR
        WGRID.text = car_menor!car_numdoc
        WGRID.Col = 6
        WGRID.CellBackColor = WCOLOR
        WGRID.text = car_menor!car_int_8dias
        WGRID.Col = 7
        WGRID.CellBackColor = WCOLOR
        WGRID.text = car_menor!car_numser
        WGRID.Col = 8
        WGRID.CellBackColor = WCOLOR
        WGRID.text = car_menor!car_numfac
        WGRID.Col = 9
        WGRID.CellBackColor = WCOLOR
        WGRID.text = Nulo_Valors(car_menor!car_concepto)
        WGRID.Rows = WGRID.Rows + 1
    End If
    car_menor.MoveNext
Loop
'WGRID.Rows = WGRID.Rows + 1
'CUENTA = CUENTA + 1
'WGRID.Row = CUENTA
'WGRID.Col = 0
'WGRID.text = "TOTAL: " & CUENTA + 1

End Sub
Public Sub LLENA_HISTORIA(wCODI As String, wCP As String, wTIPDOC As String, wFECHA_ING As String, wNUMDOC As String, WGRID As MSFlexGrid)
Dim CUENTA As Integer
Dim FLAG As Integer
Dim NUMERO As String * 3
PSCAA_HISTO.rdoParameters(0) = wCP
PSCAA_HISTO.rdoParameters(1) = Val(wCODI)
PSCAA_HISTO.rdoParameters(2) = wTIPDOC
PSCAA_HISTO.rdoParameters(3) = wFECHA_ING
PSCAA_HISTO.rdoParameters(4) = Val(wNUMDOC)
caa_histo.Requery
If caa_histo.EOF Then
    WGRID.Col = 0
    WGRID.Row = 1
    WGRID.text = "No hay Letras.."
    Exit Sub
End If
CUENTA = 0
WGRID.Rows = 2
CABE_HISTORIA
Do Until caa_histo.EOF
        CUENTA = CUENTA + 1
        WGRID.Row = CUENTA
        WGRID.Col = 0
        WGRID.text = caa_histo!CAA_FECHA_MOV
        WGRID.Col = 1
        WGRID.text = Trim(caa_histo!CAA_TIPDOC)
        WGRID.Col = 2
        WGRID.text = caa_histo!CAA_NUMDOC
        WGRID.Col = 3
        WGRID.text = caa_histo!CAA_IMPORTE
        WGRID.Col = 4
        WGRID.text = caa_histo!CAA_SALDO
        WGRID.Col = 5
        WGRID.text = caa_histo!CAA_FECHA_ING
        WGRID.Col = 6
        WGRID.text = caa_histo!CAA_fecha_vcto
        WGRID.Col = 7
        WGRID.text = Nulo_Valors(caa_histo!CAA_DESCRIPCION)
        WGRID.Col = 8
        WGRID.text = Nulo_Valors(caa_histo!CAA_CONCEPTO)
        
        WGRID.Rows = WGRID.Rows + 1
       caa_histo.MoveNext
Loop

End Sub

Public Function ALINEA2(VAR As String) As String
Dim TEMP As String * 15
Dim N1 As Integer
Dim N2 As Integer
N1 = InStr(1, VAR, " ") - 1
N2 = 3 - N1
VAR = String(N2, "    ") + Left(VAR, N1)
ALINEA2 = VAR
End Function


Public Sub CABE_LETRAS()
GridLetras.Row = 0
GridLetras.Col = 0
GridLetras.text = "Fec.Vcto."
GridLetras.Col = 1
GridLetras.text = "L.F"
GridLetras.Col = 2
GridLetras.text = "Importe "
GridLetras.Col = 3
GridLetras.text = "Situación"
GridLetras.Col = 4
GridLetras.text = "Fec.Ing."
GridLetras.Col = 5
GridLetras.text = "NUM.DOC."
GridLetras.Col = 6
GridLetras.text = "8 Dias"
GridLetras.Col = 7
GridLetras.text = "NUM.SER."
GridLetras.Col = 8
GridLetras.text = "NUM.FAC."
GridLetras.Col = 9
GridLetras.text = "Concepto"

End Sub

Public Sub CABE_HISTORIA()
GridHistoria.Row = 0
GridHistoria.Col = 0
GridHistoria.text = "Fec.MOVI."
GridHistoria.Col = 1
GridHistoria.text = "L.F"
GridHistoria.Col = 2
GridHistoria.text = "N. DOC"
GridHistoria.Col = 3
GridHistoria.text = "Importe"
GridHistoria.Col = 4
GridHistoria.text = "Saldo"
GridHistoria.Col = 5
GridHistoria.text = "Fec.ING"
GridHistoria.Col = 6
GridHistoria.text = "Fec.VCTO"
GridHistoria.Col = 7
GridHistoria.text = "Descrip."
GridHistoria.Col = 8
GridHistoria.text = "Concepto"

End Sub
