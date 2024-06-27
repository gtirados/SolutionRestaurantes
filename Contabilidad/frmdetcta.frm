VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmdetCta 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voucher"
   ClientHeight    =   6285
   ClientLeft      =   1485
   ClientTop       =   1125
   ClientWidth     =   10275
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6285
   ScaleWidth      =   10275
   Tag             =   "55"
   Begin VB.CommandButton cancelar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cancelar"
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
      Left            =   6240
      Picture         =   "frmdetcta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar PB 
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.TextBox txtglosa 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   6375
   End
   Begin VB.TextBox TEXTOVAR 
      BackColor       =   &H00FFFF00&
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "TEXTOVAR"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox TEXTOVAR2 
      BackColor       =   &H00FFFF00&
      Height          =   405
      Left            =   3240
      TabIndex        =   2
      Text            =   "TEXTOVAR2"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.ComboBox VOUCHER 
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
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   7320
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame ESTADO 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   3855
      Left            =   120
      TabIndex        =   7
      Tag             =   "100"
      Top             =   720
      Width           =   10095
      Begin VB.CommandButton cmdimp 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   5400
         Picture         =   "frmdetcta.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton salir 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Retornar"
         Height          =   555
         Left            =   6840
         Picture         =   "frmdetcta.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton grabar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3960
         Picture         =   "frmdetcta.frx":0896
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   3240
         Width           =   1095
      End
      Begin VB.ComboBox I_NUMFAC2 
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
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin ComctlLib.ListView LV_CLI 
         Height          =   375
         Left            =   3480
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   1920
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSFlexGridLib.MSFlexGrid grid_fac 
         Height          =   3015
         Left            =   90
         TabIndex        =   6
         Tag             =   "9999"
         Top             =   120
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   5318
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         FixedCols       =   0
         AllowBigSelection=   -1  'True
         FocusRect       =   2
         HighLight       =   2
         GridLines       =   2
         AllowUserResizing=   3
      End
      Begin VB.Label lblinforme 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00808000&
         Height          =   440
         Left            =   120
         TabIndex        =   33
         Top             =   3360
         Width           =   3735
      End
      Begin VB.Label lblmas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mas Informacion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Dif.:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8040
         TabIndex        =   17
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label lbldif 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   8520
         TabIndex        =   16
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label lblcodusu 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5520
         TabIndex        =   14
         Top             =   4560
         Width           =   2535
      End
      Begin VB.Label men 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   4560
         Width           =   2775
      End
      Begin VB.Label momen 
         Caption         =   "Un Momento ..."
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
         Left            =   3360
         TabIndex        =   9
         Top             =   1560
         Width           =   1575
      End
   End
   Begin ComctlLib.ProgressBar Barra 
      Height          =   255
      Left            =   240
      TabIndex        =   8
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
   Begin VB.Label lblplantilla 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4200
      TabIndex        =   34
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label lbldes 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   8880
      TabIndex        =   30
      Top             =   4920
      Visible         =   0   'False
      Width           =   1455
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
      Height          =   255
      Left            =   7560
      TabIndex        =   29
      Top             =   4800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lbl_nro_voucher 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9000
      TabIndex        =   28
      Top             =   4800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblperiodo 
      Caption         =   "PERIODO : ASIENTOS DE AJUSTES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7080
      TabIndex        =   27
      Top             =   4800
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Nro. Voucher: "
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6000
      TabIndex        =   26
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label txtnrovoucher 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   8760
      TabIndex        =   25
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "Nro de Voucher:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7200
      TabIndex        =   24
      Top             =   360
      Width           =   1410
   End
   Begin VB.Label lblglosa 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "Glosa :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblper 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "Periodo :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7200
      TabIndex        =   22
      Top             =   30
      Width           =   780
   End
   Begin VB.Label lblmes 
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
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   8040
      TabIndex        =   21
      Top             =   30
      Width           =   2055
   End
   Begin VB.Label lblplan 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "Plantilla :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3360
      TabIndex        =   20
      Top             =   0
      Width           =   810
   End
   Begin VB.Label lbllibro 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   720
      TabIndex        =   19
      Top             =   30
      Width           =   2415
   End
   Begin VB.Label lblaut 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "Libro :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   30
      Width           =   555
   End
End
Attribute VB_Name = "frmdetCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'ULTIMAS VARIABLES DEFINIDAS
'---------------------------
    Dim flag_existen_datos   As String * 1
    Dim fila_cont            As Integer
    
    
    Dim kpMOV_CODCIA         As String * 2
    Dim kpMOV_NRO_MES        As Integer
    Dim kpMOV_NRO_VOUCHER    As Integer
    Dim kpMOV_NRO_MOV        As Integer
    Dim kpMOV_TIPMOV         As Integer
    Dim kpMOV_FECHA          As Date
    Dim kpMOV_CODCTA         As String * 12
    Dim kpMOV_DH             As String * 1
    Dim kpMOV_IMPORTE        As Currency
    Dim kpMOV_GLOSA          As String
    Dim kpMOV_MONEDA         As String * 1
    Dim kpMOV_SUNAT          As String * 2
    Dim kpMOV_SERIE          As Integer
    Dim kpMOV_NUMFAC         As Currency
    Dim kpMOV_CODCLIE        As Currency
    Dim kpMOV_CP             As String * 1
    Dim kpMOV_FBG            As String * 1
    Dim kpMOV_MARCA          As String * 1
    Dim kpMOV_DETALLE        As String
    Dim kpMOV_FBG_C          As String * 1
    Dim kpMOV_SERIE_C        As Currency
    Dim kpMOV_NUMFAC_C       As Currency
    Dim kpMOV_FECHA_EMI      As Date
    Dim kpMOV_PLANTILLA      As Integer
    Dim kpMOV_FLAG_TC        As String * 1
    Dim kpMOV_TIPO_CAMBIO    As Currency
    Dim kpMOV_FLAG_DES       As String * 1
    Dim kpMOV_CODUSU         As String
    Dim kpMOV_CODTRA         As Integer
    Dim kpMOV_NUMOPER        As Integer
    Dim kpMOV_NUMOPER2       As Integer
    Dim kpMOV_CC             As Integer
    Dim kpMOV_OPC            As Currency
    Dim kpMOV_EXONERADO      As String * 1
    Dim kpMOV_FECHA_CONTABLE As Date

'---------------------------

Dim WPASA As Boolean
Dim WSELE As String * 1
Dim F2 As Integer
Dim llave1
Public WW_CUENTA As String
Dim FLAG_SALIR As Integer
Dim loc_key As Long
Dim XX_CUENTA As String * 12
Dim FILA As Integer
Dim ws_bruto_d, ws_bruto_h As Currency
Dim SUM_D As Currency
Dim SUM_H As Currency
Dim PSTEMP_LLAVE As rdoQuery
Dim temp_llave As rdoResultset
Dim LOC_ITEM As Integer
Dim cop_llave As rdoResultset
Dim PSCOP_LLAVE As rdoQuery
Dim LOC_CANCELA As Integer
Dim PSTEMP_MAYOR As rdoQuery
Dim temp_mayor As rdoResultset
Dim PSMOV_LLAVE As rdoQuery
Dim mov_llave As rdoResultset
Public ws_nro_voucher As Long
Public WS_NRO_MOV As Long
Dim PSMOV_VOU As rdoQuery
Dim VOU_MOV As rdoResultset
Dim leer_docu As rdoResultset
Dim PSDOCU As rdoQuery

Dim PSMOV_FAC As rdoQuery
Dim FAC_MOV As rdoResultset

Dim temporal
Dim wfila_act As Integer
Dim loc_ini As Integer
Dim loc_fin  As Integer
Dim Wsec As Integer
Dim loc_voucher As Integer
Dim FLAG_DIF_TC As String * 1
Dim LOC_DIF_TC As Currency
Dim CUENTA_DIF As String * 12
Dim LOC_CODCLIE As Currency
Dim loc_cp As String * 1
Dim LOC_DOCU As String

Option Explicit
Public Sub grabar_movicont()

    temp_llave.AddNew
    temp_llave!MOV_CODCIA = kpMOV_CODCIA
    temp_llave!MOV_NRO_MES = kpMOV_NRO_MES
    temp_llave!MOV_NRO_VOUCHER = kpMOV_NRO_VOUCHER
    temp_llave!MOV_NRO_MOV = kpMOV_NRO_MOV
    temp_llave!MOV_TIPMOV = kpMOV_TIPMOV
    temp_llave!MOV_FECHA = kpMOV_FECHA_CONTABLE
    temp_llave!MOV_CODCTA = kpMOV_CODCTA
    temp_llave!MOV_DH = kpMOV_DH
    temp_llave!MOV_IMPORTE = kpMOV_IMPORTE
    temp_llave!MOV_GLOSA = kpMOV_GLOSA
    temp_llave!MOV_MONEDA = kpMOV_MONEDA
    temp_llave!MOV_SUNAT = kpMOV_SUNAT
    temp_llave!MOV_SERIE = kpMOV_SERIE
    temp_llave!MOV_numfac = kpMOV_NUMFAC
    temp_llave!MOV_CODCLIE = kpMOV_CODCLIE
    temp_llave!MOV_CP = kpMOV_CP
    temp_llave!MOV_FBG = kpMOV_FBG
    temp_llave!MOV_MARCA = kpMOV_MARCA
    temp_llave!MOV_DETALLE = kpMOV_DETALLE
    temp_llave!MOV_FBG_C = kpMOV_FBG_C
    temp_llave!MOV_SERIE_C = kpMOV_SERIE_C
    temp_llave!MOV_NUMFAC_C = kpMOV_NUMFAC_C
    temp_llave!MOV_fecha_EMI = kpMOV_FECHA_EMI
    temp_llave!MOV_PLANTILLA = kpMOV_PLANTILLA
    temp_llave!MOV_FLAG_TC = kpMOV_FLAG_TC
    temp_llave!MOV_TIPO_CAMBIO = kpMOV_TIPO_CAMBIO
    temp_llave!MOV_FLAG_DES = kpMOV_FLAG_DES
    temp_llave!MOV_CODUSU = kpMOV_CODUSU
    temp_llave!MOV_CODTRA = kpMOV_CODTRA
    temp_llave!MOV_NUMOPER = kpMOV_NUMOPER
    temp_llave!MOV_NUMOPER2 = kpMOV_NUMOPER2
    temp_llave!MOV_CC = kpMOV_CC
    temp_llave!MOV_OPC = kpMOV_OPC
    temp_llave!MOV_EXONERADO = kpMOV_EXONERADO
    temp_llave!MOV_FECHA_PROCESO = kpMOV_FECHA
    temp_llave.Update
End Sub

Private Sub cancelar_Click()
'grid1.Visible = False
loc_voucher = -1

men.Caption = ""
'ESTADO.Caption = "Estado : "
FILA = 0
SUM_D = 0
SUM_H = 0
LIMPIA_DATOS
CABE_ING
frmdetCta.Lcuenta.Caption = ""
lbldif.Caption = ""
'cmdIngreso.Visible = False
'cmdIngreso.Enabled = True

SendKeys "%{UP}", True
TEXTOVAR.Visible = False
TEXTOVAR2.Visible = False
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
FILA = 0
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
Exit Sub
SALE:
MsgBox "Verficar Importe.", 48, Pub_Titulo
If TEXTOVAR.Visible Then Azul TEXTOVAR, TEXTOVAR

End Sub


Private Sub cmdcorta_Click()
LOC_CANCELA = 2
End Sub


Private Sub cmdimp_Click()
On Error GoTo SALE
Dim CADENITA
Dim wvoucher As Integer
Dim WTIPMOV As Integer
'wtipmov = Val(Right(VOUCHER.Text, 8))
wvoucher = Val(lbl_nro_voucher.Caption)
Reportes.ReportFileName = Trim(Left(PUB_RUTA_OTRO, 1)) + ":\ADMIN\CONTABILIDAD\" & "IMPVOU.RPT"
Reportes.Connect = PUB_ODBC
Reportes.WindowTitle = "Reporte :  VOUCHER "
Reportes.Destination = crptToWindow
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390
DoEvents
pub_cadena = "{MOVICONT.MOV_CODCIA} = '" & LK_CODCIA & "' and {MOVICONT.MOV_TIPMOV} = " & WTIPMOV & "  and {MOVICONT.MOV_NRO_MES} = " & 0 & " AND {MOVICONT.MOV_NRO_VOUCHER}  = " & wvoucher
Reportes.SelectionFormula = pub_cadena
Reportes.Action = 1

Exit Sub
SALE:
  MsgBox Err.Description & " // Intente Nuevamente...", 48, Pub_Titulo
End Sub



Private Sub cmdIngreso_Click()
Dim ws_tot_debe, ws_tot_haber As Currency
Dim er As rdoError
Dim pub_mensaje As String
Const ingre = 2
Const MODIF = 1
Dim N As Integer
Dim LOC_SALDO_CAR As Currency
Dim FLAG As Boolean
Dim pub_mensaje_err As String
Dim w_dh  As String

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
      If tab_mayor!tab_numtab = Val(Format(LK_FECHA_DIA, "yyyy")) Then
        wcadena = Trim(tab_mayor!tab_nomlargo)
      End If
      tab_mayor.MoveNext
    Loop
    wvalor = Mid(wcadena, 0 + 1, 1)
    If wvalor = "1" Then
       MsgBox "<<<< Mes CERRADO Operaciones >>>>", vbCritical, Pub_Titulo
       Exit Sub
    End If
End If



OTRO:

End Sub



Private Sub CmdOrden_Click()
Dim wfila As Integer
Dim wvoucher As Integer
pub_cadena = "SELECT MOV_VOU2, MOV_NRO_VOUCHER FROM MOVICONT WHERE MOV_CODCIA = ? AND (MOV_FECHA >=? AND MOV_FECHA <=?)  AND MOV_TIPMOV = ? AND MOV_NRO_MES = " & 0 & "   ORDER BY MOV_TIPMOV, MOV_NRO_VOUCHER"
Set PSTEMP_LLAVE = CN.CreateQuery("", pub_cadena)
PSTEMP_LLAVE(0) = 0
PSTEMP_LLAVE(1) = LK_FECHA_DIA
PSTEMP_LLAVE(2) = LK_FECHA_DIA
PSTEMP_LLAVE(3) = 0
Set temp_llave = PSTEMP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
PSTEMP_LLAVE(0) = LK_CODCIA
PSTEMP_LLAVE(1) = LK_FECHA_DIA
PSTEMP_LLAVE(2) = LK_FECHA_DIA
PSTEMP_LLAVE(3) = 0
temp_llave.Requery
If temp_llave.EOF Then
 MsgBox "no hay datos"
 Exit Sub
End If
wfila = 1
wvoucher = temp_llave!MOV_NRO_VOUCHER
Do Until temp_llave.EOF
If wvoucher <> temp_llave!MOV_NRO_VOUCHER Then
 wfila = wfila + 1
 wvoucher = temp_llave!MOV_NRO_VOUCHER
End If
 DoEvents

 temp_llave.Edit
 temp_llave!mov_vou2 = wfila
 temp_llave.Update
temp_llave.MoveNext
Loop




End Sub



Private Sub Form_DblClick()
Exit Sub

pub_cadena = "SELECT * FROM MOVICONT WHERE MOV_CODCIA = ?  AND MOV_NRO_MES = " & 0 & "   ORDER BY MOV_TIPMOV ,MOV_NRO_VOUCHER, MOV_DH, MOV_NRO_MOV"
Set PSTEMP_LLAVE = CN.CreateQuery("", pub_cadena)
PSTEMP_LLAVE(0) = 0
Set temp_llave = PSTEMP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
PSTEMP_LLAVE(0) = LK_CODCIA
temp_llave.Requery
If temp_llave.EOF Then
 MsgBox "no hay"
 Exit Sub
End If
Dim wv  As Integer
Dim wsuma_9 As Currency
Dim wsuma_79 As Currency
Dim WTIPMOV As Currency
wsuma_79 = 0
wsuma_9 = 0
wv = temp_llave!MOV_NRO_VOUCHER
WTIPMOV = 0
Do Until temp_llave.EOF
If temp_llave!MOV_NRO_VOUCHER <> wv Then
   If wsuma_79 <> wsuma_9 Then
     MsgBox "muestra " & wv & "   " & WTIPMOV
   End If
   wv = temp_llave!MOV_NRO_VOUCHER
   WTIPMOV = temp_llave!MOV_TIPMOV
   wsuma_79 = 0
   wsuma_9 = 0
End If
    If Left(Trim(temp_llave!MOV_CODCTA), 2) = "79" Then
     wsuma_79 = wsuma_79 + temp_llave!MOV_IMPORTE
    End If
    If Left(Trim(temp_llave!MOV_CODCTA), 1) = "9" Then
     wsuma_9 = wsuma_9 + temp_llave!MOV_IMPORTE
    End If

temp_llave.MoveNext
Loop

MsgBox "TERMINAO"
End Sub

Private Sub Form_Load()
'On Error GoTo SALE
frmdetCta.Height = 4770
frmdetCta.Top = 3195
frmdetCta.Width = 10305
frmdetCta.Left = 45


Wsec = 0
LOC_CANCELA = 0
FILA = 0
wfila_act = 0
WSELE = ""
Dim ws_indice As Integer
Dim cade

'pub_cadena = "SELECT MOV_FBG, MOV_NUMFAC, MOV_NRO_MES, MOV_NRO_VOUCHER, MOV_MONEDA, MOV_FECHA_EMI, MOV_SERIE, MOV_IMPORTE, MOV_DH , MOV_TIPMOV FROM MOVICONT WHERE MOV_CODCIA = ? AND MOV_CODCLIE = ? AND MOV_CP = ? AND MOV_FBG =  ? AND MOV_SERIE = ? AND MOV_NUMFAC = ? ORDER BY  MOV_FBG, MOV_NUMFAC"
'Set PSDOCU = CN.CreateQuery("", pub_cadena)
'PSDOCU(0) = 0
'PSDOCU(1) = 0
'PSDOCU(2) = 0
'PSDOCU(3) = 0
'PSDOCU(4) = 0
'PSDOCU(5) = 0
'Set leer_docu = PSDOCU.OpenResultset(rdOpenKeyset, rdConcurValues)


pub_cadena = "SELECT * FROM PLANTILLA WHERE PLT_CODCIA = ? AND PLT_TIPMOV = ? AND PLT_NUMERO >= ? ORDER BY  PLT_NUMERO, PLT_SECUENCIA"
Set PSTEMP_MAYOR = CN.CreateQuery("", pub_cadena)
PSTEMP_MAYOR(0) = 0
PSTEMP_MAYOR(1) = 0
PSTEMP_MAYOR(2) = 0
Set temp_mayor = PSTEMP_MAYOR.OpenResultset(rdOpenForwardOnly, rdConcurValues)


pub_cadena = "SELECT * FROM MOVICONT WHERE MOV_CODCIA = ? AND (MOV_FECHA >=? AND MOV_FECHA <=?)  AND MOV_NRO_MES = " & 0 & "   ORDER BY MOV_FECHA_EMI, MOV_NRO_VOUCHER, MOV_DH, MOV_NRO_MOV"
Set PSTEMP_LLAVE = CN.CreateQuery("", pub_cadena)
PSTEMP_LLAVE(0) = 0
PSTEMP_LLAVE(1) = LK_FECHA_DIA
PSTEMP_LLAVE(2) = LK_FECHA_DIA
Set temp_llave = PSTEMP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT MOV_FBG, MOV_NUMFAC, MOV_NRO_MES, MOV_NRO_VOUCHER, MOV_MONEDA, MOV_FECHA_EMI, MOV_SERIE, MOV_IMPORTE, MOV_DH , MOV_TIPMOV FROM MOVICONT WHERE MOV_CODCIA = ? AND MOV_CODCLIE = ? AND MOV_CP = ? AND MOV_NUMFAC <> 0 ORDER BY  MOV_FBG, MOV_SERIE , MOV_NUMFAC"
'pub_cadena = "SELECT MOV_CODCLIE, MOV_FBG, MOV_SERIE, MOV_NUMFAC FROM MOVICONT WHERE MOV_CODCIA = ? AND MOV_CODCLIE = ? AND MOV_CP = ? AND MOV_NUMFAC <> 0 GROUP BY MOV_CODCLIE, MOV_FBG, MOV_SERIE, MOV_NUMFAC"
Set PSMOV_FAC = CN.CreateQuery("", pub_cadena)
PSMOV_FAC(0) = 0
PSMOV_FAC(1) = 0
PSMOV_FAC(2) = 0
Set FAC_MOV = PSMOV_FAC.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM COPARAM WHERE COP_CODCIA = ?"
Set PSCOP_LLAVE = CN.CreateQuery("", pub_cadena)
PSCOP_LLAVE(0) = 0
Set cop_llave = PSCOP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM MOVICONT WHERE MOV_CODCIA = ?  AND MOV_NRO_VOUCHER =?  AND MOV_TIPMOV = ?  AND (MOV_FECHA >= ? AND MOV_FECHA <= ?) AND MOV_NRO_MES = " & 0 ' & " ORDER BY MOV_NRO_MOV"
Set PSMOV_LLAVE = CN.CreateQuery("", pub_cadena)
PSMOV_LLAVE(0) = 0
PSMOV_LLAVE(1) = 0
PSMOV_LLAVE(2) = 0
PSMOV_LLAVE(3) = LK_FECHA_DIA
PSMOV_LLAVE(4) = LK_FECHA_DIA
Set mov_llave = PSMOV_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT MOV_NRO_VOUCHER  FROM MOVICONT WHERE MOV_CODCIA = ? AND (MOV_FECHA >= ? AND MOV_FECHA <=?) AND MOV_NRO_MES = ? AND MOV_TIPMOV = ?   ORDER BY MOV_NRO_VOUCHER DESC "
Set PSMOV_VOU = CN.CreateQuery("", pub_cadena)
PSMOV_VOU.MaxRows = 1
PSMOV_VOU(0) = 0
PSMOV_VOU(1) = LK_FECHA_DIA
PSMOV_VOU(2) = LK_FECHA_DIA
PSMOV_VOU(3) = 0
PSMOV_VOU(4) = 0
Set VOU_MOV = PSMOV_VOU.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)

PSCOP_LLAVE.rdoParameters(0) = LK_CODCIA
cop_llave.Requery
If cop_llave.EOF Then
 Screen.MousePointer = 0
 MsgBox "Hay que definir parametros de contabilidad.", 48, Pub_Titulo
' Unload frmdetCta
 Exit Sub
End If

PUB_TIPREG = 50
PUB_CODCIA = "00"
SQ_OPER = 2
LEER_TAB_LLAVE

PUB_TIPREG = 150
PUB_CODCIA = "00"
SQ_OPER = 2
LEER_TAB_LLAVE
FLAG_DIF_TC = ""
FILA = 0
DoEvents
LIMPIA_DATOS

Dim ws_fecha As Date

FILA = 0

cancelar_Click
'cmdIngreso.Visible = False
loc_voucher = -1
' PROCESO DE MUESTRA
MUESTRA_DATA


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


Private Sub grabar_Click()
Dim kpcuenta As Integer
' CONSISTENCIAS DE INFORMACION
'******************************
If Val(grid_fac.TextMatrix(1, 11)) <> Val(grid_fac.TextMatrix(1, 12)) Then
   MsgBox "Existe diferencia en Voucher, NO CUADRA Verificar.", 48, Pub_Titulo
   Exit Sub
End If
If Val(grid_fac.TextMatrix(1, 11)) = 0 And Val(grid_fac.TextMatrix(1, 12)) = 0 Then
   MsgBox "Voucher No Tiene Importes Verificar.", 48, Pub_Titulo
   Exit Sub
End If

For fila_cont = 2 To grid_fac.Rows - 1
   If Trim(grid_fac.TextMatrix(fila_cont, 1)) = "" Then GoTo salta_consis
    SQ_OPER = 1
    PUB_CUENTA = Trim(grid_fac.TextMatrix(fila_cont, 1))
    LEER_COM_LLAVE
    If com_llave.EOF Then
       MsgBox "Cuenta No Existe.", 48, Pub_Titulo
       GoTo fin_chequeo
    Else
      'If Val(com_llave!COM_NIVEL) <> LK_NIVEL_MAX Then
      '     MsgBox "Cuenta No es Analitica.", 48, Pub_Titulo
      '     GoTo fin_chequeo
      'End If
    End If
   
    If Val(grid_fac.TextMatrix(fila_cont, 3)) <> 0 Then
     SQ_OPER = 1
     pu_cp = Trim(grid_fac.TextMatrix(fila_cont, 4))
     If pu_cp = "E" Then
       PSPER_BUSCA(0) = Val(grid_fac.TextMatrix(fila_cont, 3))
       per_busca.Requery
       If per_busca.EOF Then
           MsgBox "Codigo de " & Trim(grid_fac.TextMatrix(1, 3)) & " NO EXISTE.", 48, Pub_Titulo
           GoTo fin_chequeo
       End If
     Else
       pu_codclie = Val(grid_fac.TextMatrix(fila_cont, 3))
       pu_codcia = LK_CODCIA
       LEER_CLI_LLAVE
       If cli_llave.EOF Then
           MsgBox "Codigo de " & Trim(grid_fac.TextMatrix(1, 3)) & " NO EXISTE.", 48, Pub_Titulo
           GoTo fin_chequeo
       End If
     End If
    End If
    If Val(grid_fac.TextMatrix(fila_cont, 10)) = 1 Or Val(grid_fac.TextMatrix(fila_cont, 10)) = 2 Then
    Else
       MsgBox "Codigo no Permitido, solo 1=Grabado , 2=Exonerado", 48, Pub_Titulo
       GoTo fin_chequeo
    End If
    If Val(grid_fac.TextMatrix(fila_cont, 6)) <> 0 Then
     kpMOV_CC = Val(grid_fac.TextMatrix(fila_cont, 6))
     SQ_OPER = 1
     PUB_CODCIA = LK_CODCIA
     PUB_TIPREG = 40
     PUB_NUMTAB = Val(grid_fac.TextMatrix(fila_cont, 6))
     LEER_TAB_LLAVE
     If tab_llave.EOF Then
       MsgBox "Codigo de Centro de Costo No Existe.", 48, Pub_Titulo
       GoTo fin_chequeo
     End If
    End If
salta_consis:
Next fila_cont


If flag_existen_datos = "A" Then
  kpMOV_NRO_VOUCHER = Val(txtnrovoucher.Caption)
Else
 PSMOV_VOU(0) = LK_CODCIA
 PSMOV_VOU(1) = cop_llave!cop_fecha_proceso
 PSMOV_VOU(2) = cop_llave!cop_fecha_proceso2
 PSMOV_VOU(3) = cop_llave!COP_NRO_MES
 PSMOV_VOU(4) = Val(Trim(Right(lbllibro.Caption, 8)))
 VOU_MOV.Requery
 kpMOV_NRO_VOUCHER = 0
 If VOU_MOV.EOF Then
   kpMOV_NRO_VOUCHER = 1
 Else
   kpMOV_NRO_VOUCHER = Val(VOU_MOV!MOV_NRO_VOUCHER) + 1
 End If
End If


kpMOV_CODCIA = LK_CODCIA
kpMOV_NRO_MES = cop_llave!COP_NRO_MES
kpMOV_TIPMOV = Val(Trim(Right(lbllibro.Caption, 8)))
kpMOV_FECHA = LK_FECHA_DIA
kpMOV_GLOSA = Trim(txtglosa.Text)
'kpMOV_MONEDA = PUB_CONTAB_MONEDA
kpMOV_DETALLE = Trim(txtglosa.Text)
kpMOV_FECHA_EMI = PUB_CONTAB_FECHA_EMI
kpMOV_PLANTILLA = Val(Trim(Right(lblplantilla, 8)))
kpMOV_CODUSU = LK_CODUSU
kpMOV_CODTRA = LK_CODTRA
kpMOV_NUMOPER = PUB_CONTAB_NUMOPER
kpMOV_NUMOPER2 = PUB_CONTAB_NUMOPER2
kpMOV_FECHA_CONTABLE = cop_llave!cop_fecha_proceso2
' ELIMINAR DATOS DE ESTE VOUCHER
'********************************
pub_cadena = "DELETE MOVICONT  WHERE " & _
" MOV_CODCIA = '" & kpMOV_CODCIA & "' AND  " & _
" (MOV_FECHA >=  '" & Format(kpMOV_FECHA_CONTABLE, "yyyy/mm/dd") & "' AND MOV_FECHA <=  '" & Format(kpMOV_FECHA_CONTABLE, "yyyy/mm/dd") & "')" & _
" AND MOV_NRO_MES = " & kpMOV_NRO_MES & " AND " & _
" MOV_NRO_VOUCHER = " & kpMOV_NRO_VOUCHER & " AND MOV_TIPMOV = " & kpMOV_TIPMOV
CN.Execute pub_cadena, rdExecDirect
'********************************

kpcuenta = 0
For fila_cont = 2 To grid_fac.Rows - 1
    If Trim(grid_fac.TextMatrix(fila_cont, 1)) = "" Then GoTo SALTA_CUENTA
    
    kpMOV_NRO_MOV = kpcuenta
    kpMOV_CODCTA = Trim(grid_fac.TextMatrix(fila_cont, 1))
    If Val(grid_fac.TextMatrix(fila_cont, 11)) <> 0 Then
      kpMOV_DH = "D"
      kpMOV_IMPORTE = Val(grid_fac.TextMatrix(fila_cont, 11))
    Else
      kpMOV_DH = "H"
      kpMOV_IMPORTE = Val(grid_fac.TextMatrix(fila_cont, 12))
    End If
    kpMOV_SUNAT = Val(grid_fac.TextMatrix(fila_cont, 7))
    kpMOV_SERIE = Val(grid_fac.TextMatrix(fila_cont, 8))
    kpMOV_NUMFAC = Val(grid_fac.TextMatrix(fila_cont, 9))
    kpMOV_CODCLIE = Val(grid_fac.TextMatrix(fila_cont, 3))
    kpMOV_CP = Trim(grid_fac.TextMatrix(fila_cont, 4))
    kpMOV_FBG = " " 'grid_fac.TextMatrix(fila_cont, 1)
    kpMOV_MARCA = " " ' grid_fac.TextMatrix(fila_cont, 1)
    kpMOV_FBG_C = " " ' grid_fac.TextMatrix(fila_cont, 1)
    kpMOV_SERIE_C = 0 ' grid_fac.TextMatrix(fila_cont, 1)
    kpMOV_NUMFAC_C = 0 ' grid_fac.TextMatrix(fila_cont, 1)
    kpMOV_FLAG_TC = " " 'grid_fac.TextMatrix(fila_cont, 1)
    kpMOV_TIPO_CAMBIO = 0 ' grid_fac.TextMatrix(fila_cont, 1)
    kpMOV_FLAG_DES = " " ' grid_fac.TextMatrix(fila_cont, 1)
    kpMOV_CC = Val(grid_fac.TextMatrix(fila_cont, 6))
    kpMOV_OPC = Val(grid_fac.TextMatrix(fila_cont, 5))
    kpMOV_EXONERADO = Val(grid_fac.TextMatrix(fila_cont, 10))
    '**************
     grabar_movicont
    '**************
SALTA_CUENTA:
kpcuenta = kpcuenta + 1

Next fila_cont
MsgBox "Asiento grabado satisfactoriamente", 48, Pub_Titulo
Unload frmdetCta

Exit Sub
fin_chequeo:
If grid_fac.Visible And grid_fac.Enabled Then grid_fac.SetFocus
End Sub



Private Sub grid1_Click()

End Sub


Private Sub i_fecha2_Change()
 
End Sub

Private Sub i_fecha2_GotFocus()
If grid_fac.Visible = False Then Exit Sub

'CABE_MAN

End Sub

Private Sub i_fecha2_KeyPress(KeyAscii As Integer)
'On Error GoTo pasa
Dim WFLAG As String * 1
Dim wsFECHA1, WS_FECHA2
If KeyAscii <> 13 Then Exit Sub



End Sub


Private Sub I_NUMFAC2_Click()
  If grid_fac.Col = 8 Then
   grid_fac.CellAlignment = 1
   grid_fac.Text = Trim(I_NUMFAC2.Text)
  End If
End Sub

Private Sub I_NUMFAC2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  I_NUMFAC2.Visible = False
  grid_fac.SetFocus
  Exit Sub
End If
If KeyCode <> 13 Then Exit Sub

If grid_fac.Col = 8 Then
   grid_fac.CellAlignment = 1
   I_NUMFAC2.Visible = False
   grid_fac.Text = Trim(I_NUMFAC2.Text)
   grid_fac.SetFocus
 End If

End Sub

Private Sub I_NUMFAC2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If grid_fac.Col = 8 Then
   grid_fac.CellAlignment = 1
   grid_fac.Text = Trim(I_NUMFAC2.Text)
   grid_fac.SetFocus
 End If
End If
End Sub



Private Sub lbl_nro_voucher_DblClick()
Dim wvou
pub_mensaje = "Si desea Modificar en Nro. Correlativo del Voucher seleccione <Si>, de lo contrario siguirá el correlativo <No>"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
no_sabe:
    PSMOV_VOU.rdoParameters(0) = LK_CODCIA
    PSMOV_VOU.rdoParameters(1) = LK_FECHA_DIA
    PSMOV_VOU.rdoParameters(2) = LK_FECHA_DIA
    PSMOV_VOU.rdoParameters(3) = 0
    PSMOV_VOU.rdoParameters(4) = 0
    VOU_MOV.Requery
    If VOU_MOV.EOF Then
    ws_nro_voucher = 0
    Else
    ws_nro_voucher = VOU_MOV!MOV_NRO_VOUCHER
    End If
    ws_nro_voucher = ws_nro_voucher + 1
    lbl_nro_voucher.Caption = Format(ws_nro_voucher, "########0.0")
    loc_voucher = -1
    Exit Sub
End If
wvou = InputBox("Cambio de Nro de Voucher." & Chr(13) & "Ingrese Nro. de Voucher :", "Cambiar Voucher...", Format(lbl_nro_voucher.Caption, "######0.0"))
If wvou = "" Then
  GoTo no_sabe
End If
If Val(wvou) <= 0 Then
    MsgBox "No Procede...", 48, Pub_Titulo
    GoTo no_sabe
End If

lbl_nro_voucher.Caption = wvou
loc_voucher = Format(wvou, "#####")
'cmdIngreso.SetFocus
End Sub



Private Sub LV_CLI_DblClick()
If grid_fac.Col = 5 Then
   loc_key = LV_CLI.SelectedItem.Index
   TEXTOVAR.Text = Trim(LV_CLI.ListItems.Item(loc_key).Text) & " "
   textovar_KeyPress 13
Else
   loc_key = LV_CLI.SelectedItem.Index
   TEXTOVAR2.Text = Trim(LV_CLI.ListItems.Item(loc_key).Text) & " "
   textovar2_KeyPress 13
End If

End Sub

Private Sub LV_CLI_GotFocus()
If loc_key <> 0 Then
 Set LV_CLI.SelectedItem = LV_CLI.ListItems(loc_key)
 LV_CLI.ListItems.Item(loc_key).Selected = True
 LV_CLI.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub LV_CLI_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 If TEXTOVAR2.Visible Then
    LV_CLI.Visible = False
    TEXTOVAR2.Text = ""
    TEXTOVAR2.SetFocus
 End If
 If TEXTOVAR.Visible Then
    LV_CLI.Visible = False
    TEXTOVAR.Text = ""
    TEXTOVAR.SetFocus
 End If
End If

If KeyAscii = 13 Then
 If TEXTOVAR2.Visible Then
    textovar2_KeyPress 13
 End If
 If TEXTOVAR.Visible Then
    textovar_KeyPress 13
 End If
 
End If

End Sub

Private Sub salir_Click()
If LOC_CANCELA = 1 Then
  cmdcorta_Click
  Exit Sub
End If
Unload frmdetCta
End Sub


Public Sub LIMPIA_DATOS()

frmdetCta.grid_fac.Clear

lblcodusu.Caption = ""
FLAG_DIF_TC = ""
FILA = 1
WPASA = False
End Sub

Public Sub CABE_MOSTRAR()

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
      SUM_D = SUM_D + Val(grid_fac.TextMatrix(WF, 11))
      SUM_H = SUM_H + Val(grid_fac.TextMatrix(WF, 12))
    End If
    WF = WF + 1
    grid_fac.TextMatrix(WF - 1, 0) = Format(WF - 2, "00")
    If WF = grid_fac.Rows Then
        fx = 0
    Else
        If Trim(grid_fac.TextMatrix(WF, 1)) = "" Then fx = 0
    End If
Loop
   FILA = WF - 1
   grid_fac.TextMatrix(1, 11) = Format(SUM_D, "###,##0.00")
   grid_fac.TextMatrix(1, 12) = Format(SUM_H, "###,##0.00")
   lbldif.Caption = ""
   If (SUM_D - SUM_H) <> 0 Then lbldif.Caption = Format(SUM_D - SUM_H, "0.00")
Exit Sub
SALE:
cancelar_Click
'MsgBox "Verficar Importe.", 48, Pub_Titulo
'Resume Next
'If TEXTOVAR.Visible Then Azul TEXTOVAR, TEXTOVAR
End Sub

Private Sub Consistencias(wsGrid As MSFlexGrid, wsTexto As RichTextBox, wsKeyAscii As Integer)
  Static Valor
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
      If wsKeyAscii <> 45 And wsKeyAscii <> 8 And wsKeyAscii <> 13 And car <> "." Then
          wsKeyAscii = 0
          Beep
          Exit Sub
        End If
    End If

End Sub

Public Sub CABE_ING()
grid_fac.Cols = 23
grid_fac.Rows = 3
grid_fac.Clear
grid_fac.TextMatrix(grid_fac.Rows - 1, 10) = 1
grid_fac.RowHeight(0) = 285
grid_fac.RowHeight(1) = 285
grid_fac.RowHeight(2) = 285

FILA = 0
' falta coddigo sunat doc. relacion.

grid_fac.ColWidth(0) = 400 ' Item
grid_fac.ColWidth(1) = 800 ' Codigo Cuenta
grid_fac.ColWidth(2) = 1200 ' Descripcion de Cuenta
grid_fac.ColWidth(3) = 900 ' CODIGO DE PERSONA
grid_fac.ColWidth(4) = 0 ' SI ES C , P O E
grid_fac.ColWidth(5) = 700 ' OPC
grid_fac.ColWidth(6) = 800 ' CODIGO DE CENTRO DE COSTO
grid_fac.ColWidth(7) = 500 ' CODIGO DE SUNAT
grid_fac.ColWidth(8) = 600 ' SERIE DE DOCUMETO
grid_fac.ColWidth(9) = 800 ' NUMERO DE DOCUMENTO
grid_fac.ColWidth(10) = 400 ' FLAG DE EXONARADO
grid_fac.ColWidth(11) = 1100 ' DEBE
grid_fac.ColWidth(12) = 1100 ' HABER
grid_fac.ColWidth(13) = 0 ' CODCIA
grid_fac.ColWidth(14) = 0 ' FECHA
grid_fac.ColWidth(15) = 0 ' CODTRA
grid_fac.ColWidth(16) = 0 ' NUMOPER
grid_fac.ColWidth(17) = 0 ' NUMOPER2
grid_fac.ColWidth(18) = 0 ' si va el Cursor al DEBE o HABER
grid_fac.ColWidth(19) = 0 ' Descripcion para el Codigo
grid_fac.ColWidth(20) = 0 ' Descripcion para el OPC
grid_fac.ColWidth(21) = 0 ' Descripcion para el Centro de Costo
grid_fac.ColWidth(22) = 0 ' Descripcion para el Codigo de Sunat

grid_fac.TextMatrix(0, 0) = "It."
grid_fac.TextMatrix(0, 1) = "Cuenta"
grid_fac.TextMatrix(1, 1) = "Contable"
grid_fac.TextMatrix(0, 2) = "Descripción"
grid_fac.TextMatrix(1, 2) = "Cuenta"
grid_fac.TextMatrix(0, 3) = "Codigo"
grid_fac.TextMatrix(0, 4) = "C o P "
grid_fac.TextMatrix(0, 5) = "OPC"
grid_fac.TextMatrix(0, 6) = "Centro"
grid_fac.TextMatrix(1, 6) = "Costo"
grid_fac.TextMatrix(0, 7) = "TipDoc"
grid_fac.TextMatrix(1, 7) = "Sunat"
grid_fac.TextMatrix(0, 8) = "Serie"
grid_fac.TextMatrix(0, 9) = "Numero"
grid_fac.TextMatrix(0, 10) = "Exonerado"
grid_fac.TextMatrix(1, 10) = " (2)"
grid_fac.TextMatrix(0, 11) = "Debe"
grid_fac.TextMatrix(0, 12) = "Haber"

End Sub

Private Sub SALIR_LostFocus()
LV_CLI.Visible = False
End Sub



Public Sub BUSCAR_CTA(WTIPO As Integer)
Dim wgrupof As String
Dim wcuenta As TextBox
Dim wgrupo As String
Dim wq_cuenta As String
PUB_CODCIA = par_llave!PAR_CIACON
LK_TABLA = "CONTABLE"
If WTIPO = 1 Then
Else
 If TEXTOVAR.Text = "*" Then
  wgrupo = "" 'Trim(i_cuenta.text)
  archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_CUENTA >= '" & wgrupo & "'  ORDER BY COM_CUENTA"
 Else
 TEXTOVAR.Text = Mid(TEXTOVAR.Text, 2, Len(TEXTOVAR.Text))
 wgrupo = Trim(TEXTOVAR.Text)
 If Val(wgrupo) = 0 Then Exit Sub
 If wgrupo = "9" Then
    wgrupof = "999"
    archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_CUENTA >= '" & wgrupo & "' AND COM_CUENTA < '" & wgrupof & "'  ORDER BY COM_CUENTA"
 Else
    archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_CUENTA >= '" & wgrupo & "' AND COM_CUENTA < '" & Trim(Str(Val(wgrupo) + 1)) & "'  ORDER BY COM_CUENTA"
 End If
 End If
End If
Load frmBuscacta
frmBuscacta.lbltabla.Caption = LK_TABLA
frmBuscacta.Show 1
wq_cuenta = Trim(frmBuscacta.tcuenta)
If wq_cuenta <> "" Then
  If WTIPO = 1 Then
  Else
  TEXTOVAR2.Text = Trim(frmBuscacta.tcuenta)
  End If
End If
Unload frmBuscacta
If wq_cuenta <> "" Then
   If WTIPO = 1 Then
'     i_cuenta_KeyPress 13
   Else
     textovar2_KeyPress 13
   End If
ElseIf wq_cuenta <> "" Then
  If WTIPO = 1 Then
  Else
     textovar2_KeyPress 13
  End If
Else
  TEXTOVAR2.Visible = True
  Azul TEXTOVAR2, TEXTOVAR2
End If


End Sub



Private Sub txtglosa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  grid_fac.Col = 1
  If grid_fac.Rows >= 3 Then grid_fac.Row = 2
  grid_fac.SetFocus
End If
End Sub

Private Sub Voucher_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub


End Sub

Public Sub CABE_DOCU()
End Sub

Private Sub voucher_LostFocus()
PSMOV_VOU.rdoParameters(0) = LK_CODCIA
PSMOV_VOU.rdoParameters(1) = LK_FECHA_DIA
PSMOV_VOU.rdoParameters(2) = LK_FECHA_DIA
PSMOV_VOU.rdoParameters(3) = 0
PSMOV_VOU.rdoParameters(4) = 0
VOU_MOV.Requery
If VOU_MOV.EOF Then
   ws_nro_voucher = 0
Else
   ws_nro_voucher = VOU_MOV!MOV_NRO_VOUCHER
End If
loc_voucher = -1
ws_nro_voucher = ws_nro_voucher + 1
lbl_nro_voucher.Caption = Format(ws_nro_voucher, "########0.0")

End Sub


Public Sub MUESTRA_DATA()
Dim XS_FILA As Integer
Dim xs_cuenta As Integer

PUB_TIPREG = 150
PUB_CODCIA = "00"
PUB_NUMTAB = IRC_LIBRO
SQ_OPER = 1
LEER_TAB_LLAVE
If tab_llave.EOF Then
   MsgBox "Verificar Libro No Exiete", 48, Pub_Titulo
Else
   lbllibro.Caption = Trim(tab_llave!tab_nomlargo) & String(100, " ") & tab_llave!tab_numtab
End If

If cop_llave.EOF Then
  lblmes.Caption = "PERIODO NULO"
Else
  lblmes.Caption = Format(cop_llave!cop_fecha_proceso, "mmmm yyyy")
End If



ps_contable(0) = LK_CODCIA
ps_contable(1) = cop_llave!cop_fecha_proceso
ps_contable(2) = cop_llave!cop_fecha_proceso2
ps_contable(3) = cop_llave!COP_NRO_MES
ps_contable(4) = IRC_LIBRO
ps_contable(5) = LK_CODTRA
ps_contable(6) = PUB_CONTAB_NUMOPER2
llave_contable.Requery
flag_existen_datos = ""
If llave_contable.EOF Then
  GoSub No_data
Else
  GoSub SI_data
End If
suma_grid



Exit Sub



SI_data:

' Existe Datos o Relacion Contable
    '**************************
flag_existen_datos = "A"
fila_cont = 1
grid_fac.Rows = 2


PSTEMP_MAYOR.rdoParameters(0) = LK_CODCIA
PSTEMP_MAYOR.rdoParameters(1) = llave_contable!MOV_TIPMOV
PSTEMP_MAYOR.rdoParameters(2) = llave_contable!MOV_PLANTILLA
temp_mayor.Requery
If temp_mayor.EOF Then
Else
    If temp_mayor!PLT_SECUENCIA = 0 Then lblplantilla.Caption = Trim(temp_mayor!PLT_NOMBRE) & String(100, " ") & temp_mayor!PLT_NUMERO
End If
txtglosa.Text = Trim(llave_contable!MOV_GLOSA)
txtnrovoucher.Caption = Format(llave_contable!MOV_NRO_VOUCHER, "#####")

Do Until llave_contable.EOF
    grid_fac.Rows = grid_fac.Rows + 1
    
    fila_cont = fila_cont + 1
llave_contable.MoveNext
Loop


Return


No_data:
    ' No Existe Datos o cuando es por primera vez
    '**************************

    PSTEMP_MAYOR.rdoParameters(0) = LK_CODCIA
    PSTEMP_MAYOR.rdoParameters(1) = IRC_LIBRO
    PSTEMP_MAYOR.rdoParameters(2) = IRC_PLANTILLA
    temp_mayor.Requery
    XS_FILA = 2
    xs_cuenta = 0
    Do Until temp_mayor.EOF
       If temp_mayor!PLT_NUMERO <> IRC_PLANTILLA Then Exit Do
       If temp_mayor!PLT_SECUENCIA = 0 Then
          lblplantilla.Caption = Trim(temp_mayor!PLT_NOMBRE) & String(100, " ") & temp_mayor!PLT_NUMERO
       Else
          xs_cuenta = xs_cuenta + 1
          SQ_OPER = 1
          PUB_CUENTA = temp_mayor!PLT_CUENTA
          pu_codcia = LK_CODCIA
          PUB_CODCIA = LK_CODCIA
          LEER_COM_LLAVE
          If com_llave.EOF Then GoTo P_P
          XS_FILA = XS_FILA + 1
           'FILA = FILA + 1
           
          grid_fac.Rows = XS_FILA '+ 1
    
          grid_fac.RowHeight(grid_fac.Rows - 1) = 285
          grid_fac.TextMatrix(grid_fac.Rows - 1, 1) = Trim(temp_mayor!PLT_CUENTA)
          grid_fac.TextMatrix(grid_fac.Rows - 1, 10) = 1
          grid_fac.TextMatrix(grid_fac.Rows - 1, 4) = Trim(temp_mayor!PLT_NOMBRE)
          If Trim(temp_mayor!PLT_NOMBRE) <> "" Then ' CUANDO TIENE ALMENOS UN CODIGO DE PERSONAL.
                grid_fac.TextMatrix(grid_fac.Rows - 1, 3) = NRC_CODCLIE
                grid_fac.TextMatrix(grid_fac.Rows - 1, 4) = CRC_CP
                grid_fac.TextMatrix(grid_fac.Rows - 1, 7) = Format(CRC_SUNAT, "00")
                grid_fac.TextMatrix(grid_fac.Rows - 1, 8) = Format(NRC_NUMSER, "000")
                grid_fac.TextMatrix(grid_fac.Rows - 1, 9) = Format(NRC_NUMFAC, "0000000")
          End If
          
          If grid_fac.TextMatrix(grid_fac.Rows - 1, 4) = "C" Then
            grid_fac.TextMatrix(1, 3) = "Cliente"
          ElseIf grid_fac.TextMatrix(grid_fac.Rows - 1, 4) = "P" Then
            grid_fac.TextMatrix(1, 3) = "Proveedor"
          ElseIf grid_fac.TextMatrix(grid_fac.Rows - 1, 4) = "E" Then
            grid_fac.TextMatrix(1, 3) = "Personal"
          End If
          txtglosa.Text = CRC_GLOSA
          grid_fac.TextMatrix(grid_fac.Rows - 1, 18) = Trim(temp_mayor!PLT_DH)
          If xs_cuenta = 1 Then ' el Monto del Importe General
            If Trim(temp_mayor!PLT_DH) = "D" Then
              grid_fac.TextMatrix(grid_fac.Rows - 1, 11) = Format(NRC_IMPORTE, "0.00")
            Else
              grid_fac.TextMatrix(grid_fac.Rows - 1, 12) = Format(NRC_IMPORTE, "0.00")
            End If
          ElseIf xs_cuenta = 2 Then ' el Monto del impuesto
            If Trim(temp_mayor!PLT_DH) = "D" Then
              grid_fac.TextMatrix(grid_fac.Rows - 1, 11) = Format(NRC_IGV, "0.00")
            Else
              grid_fac.TextMatrix(grid_fac.Rows - 1, 12) = Format(NRC_IGV, "0.00")
            End If
          ElseIf xs_cuenta = 3 Then ' el Monto del del Gasto
            If Trim(temp_mayor!PLT_DH) = "D" Then
              grid_fac.TextMatrix(grid_fac.Rows - 1, 11) = Format(NRC_GASTO, "0.00")
            Else
              grid_fac.TextMatrix(grid_fac.Rows - 1, 12) = Format(NRC_GASTO, "0.00")
            End If
          End If
          
          If Not com_llave.EOF Then
            grid_fac.TextMatrix(grid_fac.Rows - 1, 2) = com_llave!com_descripcion
            'If Val(com_llave!com_nivel) <> cop_llave!cop_nivel_max Then grid_fac.TextMatrix(FILA, 2) = ""
          End If
P_P:
       End If
       temp_mayor.MoveNext
    Loop
    suma_grid
Return
End Sub
