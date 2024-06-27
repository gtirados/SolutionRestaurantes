VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F6E4F630-E903-11D5-8BB9-0080AD40A177}#1.18#0"; "OSControlsUser.ocx"
Begin VB.Form frmPaquete 
   BackColor       =   &H8000000B&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Formulacion y Produccion de Platos"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdStockProceso 
      Caption         =   "Re&gular Stock"
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
      Left            =   120
      TabIndex        =   44
      Top             =   6840
      Width           =   2430
   End
   Begin OSControlsUser.ctlText txtTotal2 
      Height          =   285
      Left            =   9240
      TabIndex        =   39
      Top             =   6210
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   503
      ForeColor       =   16711680
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImputData       =   3
      Text            =   ""
   End
   Begin OSControlsUser.ctlText txtTotal3 
      Height          =   285
      Left            =   10440
      TabIndex        =   40
      Top             =   6210
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      ForeColor       =   16711680
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImputData       =   3
      Text            =   ""
   End
   Begin OSControlsUser.OSMessage lblPaquetes 
      Height          =   330
      Left            =   0
      TabIndex        =   38
      Top             =   6000
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   582
      Picture         =   "Paquete.frx":0000
   End
   Begin VB.TextBox txtCostoPa 
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
      Height          =   285
      Left            =   8055
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   6540
      Width           =   1095
   End
   Begin VB.TextBox txtValorPa 
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
      Height          =   285
      Left            =   9255
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   6540
      Width           =   1095
   End
   Begin VB.TextBox txtUtilidadPa 
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
      Height          =   300
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   6525
      Width           =   1065
   End
   Begin VB.CommandButton cmdStock 
      Caption         =   "Agregar o &Quitar Platos"
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
      Left            =   75
      TabIndex        =   33
      Top             =   6360
      Width           =   2445
   End
   Begin OSControlsUser.OSMessage OSMessage1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   32
      Top             =   6990
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   582
      Picture         =   "Paquete.frx":04C6
   End
   Begin VB.Frame fraFind 
      BackColor       =   &H8000000B&
      Height          =   3690
      Left            =   2160
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox txtFindItem 
         Height          =   285
         Left            =   75
         TabIndex        =   12
         Top             =   270
         Width           =   5880
      End
      Begin MSComctlLib.ListView lvwItems 
         Height          =   3015
         Left            =   45
         TabIndex        =   13
         Top             =   600
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   5318
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.TextBox txtTotal1 
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
      Height          =   285
      Left            =   8055
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   6210
      Width           =   1095
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "A&ctualizar"
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
      Left            =   2730
      TabIndex        =   41
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdSetItem 
      Caption         =   "&Agregar Insumo"
      Height          =   375
      Left            =   3960
      TabIndex        =   14
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Frame fraFamilias 
      BackColor       =   &H8000000B&
      Caption         =   "&Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1740
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11835
      Begin VB.CommandButton cmdTodos 
         Caption         =   "&Listar Todos"
         Height          =   375
         Left            =   9480
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cboSubLinea 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1170
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.OptionButton optIP 
         BackColor       =   &H8000000B&
         Caption         =   "Mostrar &Platos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   0
         Left            =   7245
         TabIndex        =   22
         Top             =   690
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optIP 
         BackColor       =   &H8000000B&
         Caption         =   "&Mostrar Insumos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   1
         Left            =   7245
         TabIndex        =   21
         Top             =   1185
         Width           =   2040
      End
      Begin VB.ComboBox cboClase 
         Height          =   315
         Left            =   4890
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1170
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.ComboBox cboMarca 
         Height          =   315
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   615
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.ComboBox cboLinea 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   585
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.ComboBox cboFamilia 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1170
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.ComboBox cboDivision 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   585
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.Label lblart 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "División:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label lblart 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Familia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   930
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label lblart 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Linea:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   2625
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label lblart 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Sub Linea:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   3
         Left            =   2670
         TabIndex        =   17
         Top             =   975
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lblart 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Marca:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   4
         Left            =   4935
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label lblart 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Clase:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   5
         Left            =   4905
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   540
      End
   End
   Begin MSComctlLib.ImageList imgPaquetes 
      Left            =   3195
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   19
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Paquete.frx":0BD8
            Key             =   "item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Paquete.frx":10A4
            Key             =   "raiz"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Paquete.frx":14E8
            Key             =   "paquete"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trwPaquetes 
      Height          =   1845
      Left            =   0
      TabIndex        =   9
      Top             =   2520
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3254
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
   End
   Begin VB.CommandButton Command1 
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
      Left            =   4050
      TabIndex        =   30
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   0
      TabIndex        =   8
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Fecha de Duración del Plato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   6555
      TabIndex        =   25
      Top             =   2415
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CheckBox Check1 
         Caption         =   "Fechas de Duración "
         Height          =   240
         Left            =   2805
         TabIndex        =   42
         Top             =   240
         Width           =   2145
      End
      Begin OSControlsUser.ctlMaskEdBox mskFechaFin 
         Height          =   315
         Left            =   1545
         TabIndex        =   29
         Top             =   1110
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         Format          =   "ddddd"
      End
      Begin OSControlsUser.ctlMaskEdBox mskFechaIni 
         Height          =   315
         Left            =   1530
         TabIndex        =   28
         Top             =   645
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         Format          =   "ddddd"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final:"
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
         Left            =   165
         TabIndex        =   27
         Top             =   1110
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial:"
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
         Left            =   180
         TabIndex        =   26
         Top             =   675
         Width           =   1170
      End
   End
   Begin OSControlsUser.OSNewFlexGrid OSNewFlexGrid1 
      Height          =   1455
      Left            =   0
      TabIndex        =   23
      Top             =   4440
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   2566
      BackColor       =   16777215
      ForeColor       =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      BorderStyle     =   0
      BackColorSel    =   8388608
      Cols            =   10
      ColEdit(3)      =   0   'False
      ColSalto(3)     =   0
      ColAtras(3)     =   0
      FixedCols       =   3
      ColWidth0       =   1020
   End
   Begin VB.Label lblTitleDel 
      Caption         =   "Para Eliminar Insumo: Presione la Barra Espaciadora y se pintara de rojo.Despues presione SUPR"
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
      Left            =   4425
      TabIndex        =   43
      Top             =   5985
      Visible         =   0   'False
      Width           =   8415
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Datos del Plato :"
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
      Left            =   6000
      TabIndex        =   37
      Top             =   6585
      Width           =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Totales :"
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
      Left            =   6870
      TabIndex        =   31
      Top             =   6270
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "&Buscar Platos o Insumos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   2100
   End
End
Attribute VB_Name = "frmPaquete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'DECLARO PARA SABER CUANDO ACTULIZO EN PAQUETES
Dim vActu As Integer
Dim vmesage As Boolean

'REGISTRO DE PAQUETES O ARTICULOS
Dim rsP As New ADODB.Recordset
Dim CmdIp As New ADODB.Command


'REGISTRO DE FAMILIAS,CLASE ETC
Dim rs As New ADODB.Recordset
Dim CmdFamilias As New ADODB.Command

Dim Cnt As New ADODB.Connection
Dim Flag_Unidad As String
Dim oNodo As MSComctlLib.Node
Dim oItem As MSComctlLib.ListItem
Dim Flag_Set_Find As Byte

Dim PrmCodCia As New ADODB.Parameter
Dim PrmTipReg As New ADODB.Parameter
Dim PrmCodArt As New ADODB.Parameter

Dim PrmPack As New ADODB.Parameter
Dim PrmDiv As New ADODB.Parameter
Dim PrmFam As New ADODB.Parameter
Dim PrmLinea As New ADODB.Parameter
Dim PrmSublinea As New ADODB.Parameter
Dim PrmMarca As New ADODB.Parameter
Dim PrmClase As New ADODB.Parameter
Dim PrmDesc As New ADODB.Parameter
Dim PrmAlterno As New ADODB.Parameter
Dim StockPack As Integer


Private Sub Check1_Click()
If Check1.Value = 1 Then
 mskFechaIni.Visible = True
 mskFechaFin.Visible = True
Else
 mskFechaIni.Visible = Not True
 mskFechaFin.Visible = Not True
End If
End Sub
Private Sub cmdStock_Click()
Dim stock As Long
stock = Val(InputBox(Prompt:="Ingrese la cantidad de platos. " & _
                           "Si desea disminuir la cantidad de platos " & _
                           "ingrese el signo menos delante: Ejemplo " & _
                           "-10 ", _
                  Title:="Aumento ó disminución de platos", _
                  Default:=0))
If stock = 0 Then
 If MsgBox("La cantidad de platos que desea aumentar o " & _
        "disminuir es 0. !!!¿Desea continuar de todas maneras?¡¡¡ ", _
        vbInformation + vbYesNo, "Aumentar o Disminuir platos") = vbNo Then
   Exit Sub
 End If
End If
If Not stock = 0 Then
 vActu = 0
 UpdateStock (stock)
End If
End Sub

Private Sub cmdStockProceso_Click()
Dim i As Long
      
    For i = 2 To trwPaquetes.Nodes.count - 1
      Call trwPaquetes_NodeClick(trwPaquetes.Nodes(i))
       trwPaquetes.Nodes(i).Selected = True
      If StockPack < 0 Then
          vActu = 0
          vmesage = True
          StockPack = -(StockPack)
          UpdateStock StockPack
      End If
    Next
    vmesage = False
End Sub

Private Sub Form_Activate()
txtFind.SetFocus
OSNewFlexGrid1.Rows = 1
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With lvwItems
  .SortKey = ColumnHeader.Index - 1
  .Sorted = True
End With
End Sub

Private Sub mskFechaFin_GotFocus()
OSMessage1.Text = "Ingrese la fecha con la que se termina la venta del plato"
End Sub
Private Sub mskFechaIni_GotFocus()
OSMessage1.Text = "Ingrese la fecha con la que se inicia la venta del plato"
End Sub
Private Sub OSNewFlexGrid1_DeleteRow(Cancel As Boolean, intColumn As Integer, introw As Integer)
Dim SQL As String
On Error GoTo ErrorHandle
With OSNewFlexGrid1
 If optIP(0) = True Then
   If MsgBox("Esta seguro de eliminar el Insumo: " & vbCrLf & "" & _
   CStr(.TextMatrix(CLng(introw), 2)) & vbCrLf & "Del plato: " & vbCrLf & "" & _
   Trim(trwPaquetes.SelectedItem.Text), vbYesNo, "Eliminación") = vbYes Then
    Cancel = False
    SQL = "DELETE FROM PAQUETES " & _
    "WHERE PA_CODCIA='" & LK_CODCIA & "' AND " & _
    "PA_CODPA='" & Trim(Val(trwPaquetes.SelectedItem.Key)) & "' " & _
    "AND PA_CODART='" & .TextMatrix(CLng(introw), 0) & "'"
    Cnt.Execute SQL, , adExecuteNoRecords
    trwPaquetes.SetFocus
   Else
    Cancel = True
    trwPaquetes.SetFocus
    Exit Sub
   End If
Else
   MsgBox Prompt:="Imposible eliminar este Insumo," & _
                  "puede realizarlo desde el maestro de articulos", _
                  Buttons:=vbInformation, _
                  Title:="Eliminación de Insumos"
   Cancel = True
   trwPaquetes.SetFocus
   Exit Sub
  End If
End With
Exit Sub
ErrorHandle:
 Select Case Err.Number
  Case Is = -2147217900
   MsgBox DeleteMsg(Text:=Err.Description)
   Cancel = True
  Case Else
   MsgBox Err.Description
  End Select
End Sub
Private Sub OSNewFlexGrid1_ExitCell(introw As Integer, intColumn As Integer)
Dim SQL As String
Dim rs  As New ADODB.Recordset
Dim rsE As New ADODB.Recordset
Dim Sql2 As String
Dim var As Single
Dim sngIGV As Single
Dim i As Long
Dim C As Long
On Error GoTo ErrorHandle
Select Case intColumn
 Case Is = 5
   Sql2 = "select arm_stock from articulo " & _
   "WHERE arm_CODCIA='" & LK_CODCIA & "' AND " & _
   "arm_CODART='" & Val(OSNewFlexGrid1.TextMatrix(CLng(introw), 0)) & "'"
   rsE.Open Sql2, Cnt, adOpenForwardOnly, adCmdText
   
   SQL = "SELECT PRE_CODCIA, PRE_EQUIV,PRE_CODART, PRE_COSTO,PRE_PRE11 FROM precios " & _
   "WHERE PRE_CODCIA='" & LK_CODCIA & "' AND " & _
   "PRE_CODART='" & Val(OSNewFlexGrid1.TextMatrix(CLng(introw), 0)) & "' and " & _
   "PRE_UNIDAD='" & Trim(OSNewFlexGrid1.TextMatrix(CLng(introw), 5)) & "'"
   
  rs.Open SQL, Cnt, adOpenForwardOnly, adLockOptimistic, adCmdText
  sngIGV = (LK_IGV / 100) + 1
  With OSNewFlexGrid1
  While Not rs.EOF
    If Not Trim(Flag_Unidad) = Trim(.TextMatrix(CLng(introw), 5)) Then
     .TextMatrix(CLng(introw), 6) = Format(rsE!ARM_STOCK / (rs!PRE_EQUIV), "##,##0.0000")
      Flag_Unidad = ""
    End If
    .TextMatrix(CLng(introw), 7) = Format(rs!PRE_COSTO, "##,##0.0000")
    .TextMatrix(CLng(introw), 8) = Format(rs!pre_pre11 / sngIGV, "##,##0.0000")
    If rs!PRE_COSTO > 0 Then
    .TextMatrix(CLng(introw), 9) = Format((((rs!pre_pre11 / sngIGV) / rs!PRE_COSTO) - 1) * 100, "##,##0.0000")
    Else
    .TextMatrix(CLng(introw), 9) = "0.00"
    End If
    rs.MoveNext
  Wend
  txtTotal1 = "0.00"
  txtTotal2 = "0.00"
  txtTotal3 = "0.00"
  For C = 1 To .Rows - 1
   txtTotal1 = Format(CDbl("0" & txtTotal1) + CDbl("0" & .TextMatrix(C, 7)), "##,##0.0000")
   txtTotal2 = Format(CDbl("0" & txtTotal2) + CDbl("0" & .TextMatrix(C, 8)), "##,##0.0000")
  Next
  txtTotal3 = Format(((CDbl(txtTotal2) / CDbl(txtTotal1)) - 1) * 100, "##,##0.0000")
  End With
  
  Set rs = Nothing
  Set rsE = Nothing
  Exit Sub
 Case 4 'ACTUALIZACIÓN DE EQUIVALENCIAS
   With OSNewFlexGrid1
   If Val(.TextMatrix(CLng(introw), 4)) > 0 And Val(.TextMatrix(CLng(introw), 3)) > 0 Then
     var = Val(.TextMatrix(CLng(introw), 4)) / Val(.TextMatrix(CLng(introw), 3))
     If Val(.TextMatrix(CLng(introw), 4)) Mod Val(.TextMatrix(CLng(introw), 3)) = 0 Then
      For i = 1 To .Rows - 1
       .TextMatrix(i, 4) = Format(Val(.TextMatrix(i, 3)) * var, "##,##0.00")
      Next
     Else
      MsgBox Prompt:="Imposible realizar estas equivalencias. Las " & _
      "equivalencias ingresadas producen datos inexactos, verifique " & _
      "los datos de la columna Equiv. y reintente", _
      Buttons:=vbInformation, _
      Title:="Equivalencias"
      .SetFocus
      Exit Sub
     End If
   End If
   End With
   'SendKeys "{Esc}"
   Exit Sub
End Select
Exit Sub
ErrorHandle:
 Set rs = Nothing
 Set rsE = Nothing
 MsgBox Err.Description
 Exit Sub
End Sub
Private Sub OSNewFlexGrid1_IntroCell(introw As Integer, intColumn As Integer)
Dim SQL As String
Dim rs  As New ADODB.Recordset
On Error GoTo ErrorHandle
   Flag_Unidad = OSNewFlexGrid1.TextMatrix(CLng(introw), 5)
   
   SQL = "SELECT PRE_CODCIA, PRE_CODART, PRE_SECUENCIA,PRE_UNIDAD FROM precios " & _
   "WHERE PRE_CODCIA='" & LK_CODCIA & "' AND PRE_CODART='" & Val(OSNewFlexGrid1.TextMatrix(CLng(introw), 0)) & "'"
  
  rs.Open SQL, Cnt, adOpenForwardOnly, adLockOptimistic, adCmdText
  While Not rs.EOF
   With OSNewFlexGrid1
    .AddItemCombo CStr(rs!pre_unidad), CLng(intColumn)
    rs.MoveNext
   End With
  Wend
  Set rs = Nothing
  
  Exit Sub
ErrorHandle:
 Set rs = Nothing
 MsgBox Err.Description
 Exit Sub
End Sub
Private Sub cboClase_Click()
PrmClase = Trim(str(Val(cboClase.ItemData(cboClase.ListIndex))))
PrmDesc = Null
PrmAlterno = Null
Requery
End Sub
Private Sub cboDivision_Click()
On Error GoTo ErrorHandle
PrmTipReg = 123
PrmDiv = Trim(str(Val(cboDivision.ItemData(cboDivision.ListIndex))))
PrmCodArt = Val(cboDivision.ItemData(cboDivision.ListIndex))
Load_Familias cboFamilia
PrmFam = Null
PrmLinea = Null
PrmSublinea = Null
PrmMarca = Null
PrmClase = Null
PrmDesc = Null
PrmAlterno = Null

Requery
If Not cboDivision.ListIndex = cboDivision.ListIndex Then
 cboLinea.Clear
 cboSubLinea.Clear
 cboClase.Clear
 cboMarca.Clear
End If
ErrorHandle:
End Sub
Private Sub cboFamilia_Click()
On Error GoTo ErrorHandle
PrmTipReg = 129
PrmCodArt = Val(cboFamilia.ItemData(cboFamilia.ListIndex))
PrmFam = Trim(str(Val(cboFamilia.ItemData(cboFamilia.ListIndex))))
PrmLinea = Null
PrmSublinea = Null
PrmMarca = Null
PrmClase = Null
PrmDesc = Null
PrmAlterno = Null
Load_Familias cboLinea

Requery

cboSubLinea.Clear
cboClase.Clear
cboMarca.Clear
ErrorHandle:
End Sub
Private Sub cboLinea_Click()
PrmTipReg = 130
If Not cboLinea.ListIndex = -1 Then
PrmCodArt = Val(cboLinea.ItemData(cboLinea.ListIndex))
PrmLinea = Trim(str(Val(cboLinea.ItemData(cboLinea.ListIndex))))
End If
PrmSublinea = Null
PrmMarca = Null
PrmClase = Null
PrmDesc = Null
PrmAlterno = Null

Load_Familias cboSubLinea

Requery

cboClase.Clear
cboMarca.Clear
End Sub
Private Sub cboMarca_Click()
PrmTipReg = 132
PrmCodArt = Val(cboMarca.ItemData(cboMarca.ListIndex))

PrmMarca = Trim(str(Val(cboMarca.ItemData(cboMarca.ListIndex))))
PrmClase = Null
PrmDesc = Null
PrmAlterno = Null
Load_Familias cboClase

Requery

End Sub
Private Sub cboSubLinea_Click()
PrmTipReg = 131
PrmCodArt = Val(cboSubLinea.ItemData(cboSubLinea.ListIndex))
PrmSublinea = Trim(str(Val(cboSubLinea.ItemData(cboSubLinea.ListIndex))))
PrmMarca = Null
PrmClase = Null
PrmDesc = Null
PrmAlterno = Null
Load_Familias cboMarca
Requery
cboClase.Clear
End Sub
Private Sub cmdActualizar_Click()
 If OSNewFlexGrid1.Rows <= 1 Then '--JC
    MsgBox "No hay Datos a Actualizar", vbExclamation, Pub_Titulo
    Exit Sub
 End If
 UpdateStock
End Sub
'*****************************************************************
'PROCEDIMIENTO PARA ACTUALIZAR LOS DATOS EN LAS TABLAS PRINCIPALES
'FACART,ALLOG, ACTUALIZACION DE STOCKS, PRECIOS
'*****************************************************************
Private Sub UpdateStock(Optional ByVal C As Long = 0)
Dim SQLp As String
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim SQL4 As String
Dim i As Long
Dim NumFac As Long
Dim Equiv As Long
Dim NumOper As Long
Dim Cantidad As Long
Dim CantTotal As Long ' cantidad a descontar de la tabla paqutes
'On Error GoTo ErrorHandle
If Check1.Value = 1 Then
    If Not (CDate(mskFechaFin) > CDate(mskFechaIni) _
     And IsDate(mskFechaIni.Text) And IsDate(mskFechaFin)) Then
     MsgBox Prompt:="Las fechas de duración del plato no son validas", _
            Title:="Validación de Fechas", _
            Buttons:=vbInformation
     mskFechaIni.SetFocus
     Exit Sub
    End If
End If
'If CDbl("0" & txtTotal3) < 0 Then
' MsgBox Prompt:="Imposible guardar los datos, verifique las utilidades", _
'        Title:="Validación de Utilidades", _
'        Buttons:=vbInformation
' mskFechaIni.SetFocus
' Exit Sub
'End If
If optIP(0).Value = True Then
    NumFac = GetNextNumberFac()
    NumOper = GetNumberOper()
   If OSNewFlexGrid1.Rows > 1 Then
    With OSNewFlexGrid1
    Cnt.BeginTrans
     For i = 1 To .Rows - 1
      If C = 0 Then
       Cantidad = Val((.TextMatrix(CLng(i), 4))) / Val((.TextMatrix(CLng(i), 3)))
      Else
       Cantidad = C
      End If
      CantTotal = Val(.TextMatrix(CLng(i), 3)) * Cantidad
      If Check1.Value = 1 Then
      'SI SE ACTUALIZA LAS FECHAS
      SQL = "UPDATE PAQUETES SET " & _
            "PA_CANTIDAD=PA_CANTIDAD + '" & CantTotal & "',PA_UNIDAD='" & _
            Trim(.TextMatrix(CLng(i), 5)) & "',PA_PROM='" & _
            (.TextMatrix(CLng(i), 3)) & "',PA_FECHA_INI='" & _
             mskFechaIni & "',PA_FECHA_FIN='" & _
             mskFechaFin & "',PA_FLAG_ANULADO='S' " & _
            "WHERE PA_CODCIA='" & LK_CODCIA & "' AND " & _
            "PA_CODPA='" & Val(trwPaquetes.SelectedItem.Key) & "' AND " & _
            "PA_CODART='" & Val(.TextMatrix(CLng(i), 0)) & "'"
      Else
       SQL = "UPDATE PAQUETES SET " & _
            "PA_CANTIDAD=PA_CANTIDAD + '" & CantTotal & "',PA_UNIDAD='" & _
            Trim(.TextMatrix(CLng(i), 5)) & "',PA_PROM='" & _
            (.TextMatrix(CLng(i), 3)) & "',PA_FLAG_ANULADO='S' " & _
            "WHERE PA_CODCIA='" & LK_CODCIA & "' AND " & _
            "PA_CODPA='" & Val(trwPaquetes.SelectedItem.Key) & "' AND " & _
            "PA_CODART='" & Val(.TextMatrix(CLng(i), 0)) & "'"
     
      End If
      Debug.Print SQL
      
      Equiv = GetEquivArti((Val(.TextMatrix(CLng(i), 0))))
     
       'Cnt.Execute "exec sp_check_stock '" & LK_CODCIA & "','" & _
        OSNewFlexGrid1.TextMatrix(CLng(I), 0) & "','" & _
        Val(.TextMatrix(CLng(I), 3)) * Cantidad * Equiv & "','" & _
        Val(trwPaquetes.SelectedItem.Key) & "'", , adExecuteNoRecords
        
      'Sql3 = "INSERT INTO FACART (FAR_TIPMOV,FAR_CODCIA,FAR_NUMSER,FAR_FBG,FAR_NUMFAC,FAR_NUMSEC, " & _
             "FAR_FECHA,FAR_NUMOPER,FAR_CODART,FAR_ESTADO,FAR_SIGNO_ARM,FAR_PRECIO, " & _
             "FAR_COSPRO,FAR_EQUIV,FAR_NUMSER_C,FAR_NUMFAC_C,FAR_CANTIDAD,FAR_DESCRI, " & _
             "FAR_FECHA_PRO,FAR_FECHA_COMPRA,FAR_FLAG_SO,FAR_TRANSITO,FAR_NUMOPER2,FAR_CONCEPTO,FAR_CODUSU) " & _
             "VALUES('70','" & LK_CODCIA & "','1','','" & NumFac & "','" & _
             I & "','" & LK_FECHA_DIA & "','" & NumOper & "','" & Val(.TextMatrix(CLng(I), 0)) & "','" & _
             UCase("N") & "','" & IIf(Cantidad > 0, "-1", "+1") & "','" & CDbl("0" & .TextMatrix(CLng(I), 8)) + CDbl("0" & .TextMatrix(CLng(I), 8)) * (0.18) & "','" & _
             CDbl("0" & .TextMatrix(CLng(I), 7)) & "','" & Equiv & "','" & _
             1 & "','" & NumFac & "','" & Val(.TextMatrix(CLng(I), 3)) * Equiv * IIf(Cantidad > 0, Cantidad, Cantidad * -1) & "','" & _
             Trim((.TextMatrix(CLng(I), 5))) & "','" & LK_FECHA_DIA & "','" & _
             LK_FECHA_DIA & "','','','" & NumOper & "','" & "Transf. a Paquetes: " & Val(trwPaquetes.SelectedItem.Key) & "','" & LK_CODUSU & "')"
             
    Sql3 = "INSERT INTO FACART (FAR_TIPMOV,FAR_CODCIA,FAR_NUMSER,FAR_FBG,FAR_NUMFAC,FAR_NUMSEC, " & _
             "FAR_FECHA,FAR_NUMOPER,FAR_CODART,FAR_ESTADO,FAR_SIGNO_ARM,FAR_PRECIO, " & _
             "FAR_COSPRO,FAR_EQUIV,FAR_NUMSER_C,FAR_NUMFAC_C,FAR_CANTIDAD,FAR_DESCRI, " & _
             "FAR_FECHA_PRO,FAR_FECHA_COMPRA,FAR_FLAG_SO,FAR_TRANSITO,FAR_NUMOPER2,FAR_CONCEPTO,FAR_CODUSU) " & _
             "VALUES('70','" & LK_CODCIA & "','1','','" & NumFac & "','" & _
             i & "','" & LK_FECHA_DIA & "','" & NumOper & "','" & Val(.TextMatrix(CLng(i), 0)) & "','" & _
             UCase("N") & "','" & IIf(Cantidad > 0, "-1", "+1") & "','" & CDbl("0" & .TextMatrix(CLng(i), 8)) + CDbl("0" & .TextMatrix(CLng(i), 8)) * (0.18) & "','" & _
             CDbl("0" & .TextMatrix(CLng(i), 7)) & "','" & Equiv & "','" & _
             1 & "','" & NumFac & "','" & IIf(Cantidad > 0, Val(.TextMatrix(CLng(i), 3)) * Equiv * Cantidad, Val(.TextMatrix(CLng(i), 3)) * Equiv * Cantidad * -1) & "','" & _
             Trim((.TextMatrix(CLng(i), 5))) & "','" & LK_FECHA_DIA & "','" & _
             LK_FECHA_DIA & "','','','" & NumOper & "','" & "Transformacion en platos: " & Val(trwPaquetes.SelectedItem.Key) & "','" & LK_CODUSU & "')"
          
      If Cantidad > 0 Then
       SQLp = "UPDATE ARTICULO SET ARM_STOCK=ARM_STOCK - " & Val(.TextMatrix(CLng(i), 3)) * Cantidad * Equiv & " " & _
                  "WHERE ARM_CODCIA='" & LK_CODCIA & "' AND ARM_CODART='" & Val(.TextMatrix(CLng(i), 0)) & "'"
      Else
       SQLp = "UPDATE ARTICULO SET ARM_STOCK=ARM_STOCK + " & Val(.TextMatrix(CLng(i), 3)) * Cantidad * Equiv * -1 & " " & _
                  "WHERE ARM_CODCIA='" & LK_CODCIA & "' AND ARM_CODART='" & Val(.TextMatrix(CLng(i), 0)) & "'"
      End If
      
                  
      'Cnt.Execute SQLp, , adExecuteNoRecords
      
      'Cnt.Execute SQL + " " + Sql3, , adExecuteNoRecords
      
      If vActu = 0 Then
        Cnt.Execute SQLp, , adExecuteNoRecords
        Cnt.Execute SQL + " " + Sql3, , adExecuteNoRecords
      Else
        Cnt.Execute SQL
      End If
      
      
     Next
    End With
    
     'ACTUALIZACIÓN DEL PRECIO DEL ARTICULO EN LA TABLA PRECIOS
     If txtTotal2.Enabled = True Then
        If vActu = 0 Then
            Cnt.Execute "UPDATE PRECIOS SET PRE_PRE11='" & CDbl("0" & txtTotal2) + (CDbl("0" & txtTotal2) * (LK_IGV / 100)) & "'" & _
                  "WHERE PRE_CODCIA='" & LK_CODCIA & "' AND " & _
                  "PRE_CODART='" & Val(trwPaquetes.SelectedItem.Key) & "'", , adExecuteNoRecords
        End If
     End If
     'INSERCIÓN DEL REGISTRO EN LA TABLA FACART
     Sql3 = "INSERT INTO FACART (FAR_TIPMOV,FAR_CODCIA,FAR_NUMSER,FAR_FBG,FAR_NUMFAC,FAR_NUMSEC, " & _
             "FAR_FECHA,FAR_NUMOPER,FAR_CODART,FAR_ESTADO,FAR_SIGNO_ARM,FAR_PRECIO, " & _
             "FAR_COSPRO,FAR_EQUIV,FAR_NUMSER_C,FAR_NUMFAC_C,FAR_CANTIDAD,FAR_DESCRI," & _
             "FAR_FECHA_PRO,FAR_FECHA_COMPRA,FAR_FLAG_SO,FAR_TRANSITO,FAR_NUMOPER2,FAR_CONCEPTO,FAR_CODUSU) " & _
             "VALUES('70','" & LK_CODCIA & "','1','','" & NumFac & "','" & _
             i & "','" & LK_FECHA_DIA & "','" & NumOper & "','" & Val(trwPaquetes.SelectedItem.Key) & "','" & _
             UCase("N") & "','" & IIf(Cantidad > 0, "+1", "-1") & "','" & CDbl("0" & txtTotal2) + CDbl("0" & txtTotal2) * (0.18) & "','" & _
             CDbl("0" & txtTotal1) & "','1','" & _
             1 & "','" & NumFac & "','" & IIf(Cantidad > 0, Cantidad, Cantidad * -1) & "','UNIDAD','" & LK_FECHA_DIA & "','" & LK_FECHA_DIA & "','','','" & _
             NumOper & "','" & "Formac. de Platos: " & Val(trwPaquetes.SelectedItem.Key) & "','" & LK_CODUSU & "')"
     'INSERCIÓN DEL REGISTRO EN LA TABLA ALLOG
     SQL4 = "INSERT INTO ALLOG (ALL_CODCIA,ALL_FECHA_DIA,ALL_NUMOPER," & _
             "ALL_CODTRA,ALL_FLAG_EXT,ALL_CODART,ALL_SECUENCIA,ALL_FBG,ALL_CP," & _
             "ALL_CANTIDAD,ALL_NUMSER,ALL_NUMFAC,ALL_TIPMOV,ALL_NUMSER_C,ALL_NUMFAC_C," & _
             "ALL_FECHA_PRO,ALL_CODUSU,ALL_CONCEPTO,ALL_NUMOPER2,ALL_TIPO_CAMBIO) " & _
             "VALUES('" & LK_CODCIA & "','" & LK_FECHA_DIA & "','" & NumOper & "','2400','N','" & _
             Val(trwPaquetes.SelectedItem.Key) & "','0','','','" & _
             IIf(Cantidad > 0, Cantidad, Cantidad * -1) & "','1','" & NumFac & "','70','1','" & _
             NumFac & "','" & LK_FECHA_DIA & "','" & _
             LK_CODUSU & "','" & "Formac. de Platos: " & Val(trwPaquetes.SelectedItem.Key) & "','" & NumOper & "','" & LK_TIPO_CAMBIO & "')"
     
      'Cnt.Execute Sql3 + " " + SQL4, , adExecuteNoRecords
      If vActu = 0 Then
        Cnt.Execute Sql3 + " " + SQL4, , adExecuteNoRecords
      End If
      
     'actualización del stock del articulo paquete
     If vActu = 0 Then
        Cnt.Execute "UPDATE ARTICULO SET ARM_COSPRO='" & CDbl("0" & txtTotal1) & "',ARM_STOCK=ARM_STOCK + '" & Cantidad & "'" & _
            "WHERE ARM_CODCIA='" & LK_CODCIA & "' AND " & _
             "ARM_CODART='" & Val(trwPaquetes.SelectedItem.Key) & "'", , adExecuteNoRecords
      End If
     Cnt.CommitTrans
     
   ' MsgBox Prompt:="La actualización se completo satisfactoriamente", _
    Buttons:=vbInformation, _
    Title:="Actualización de Platos"
    Call ShowItems(Val(trwPaquetes.SelectedItem.Key), Trim(LK_CODCIA))
    Call ShowDataPack(Val(trwPaquetes.SelectedItem.Key))
    trwPaquetes.SetFocus
    Exit Sub
   End If
   trwPaquetes.SetFocus
   Exit Sub
End If
Exit Sub
ErrorHandle:
 Select Case Err.Number
  Case Is = 3021
   MsgBox Prompt:="No existe ningún insumo para este plato seleccionado", _
           Title:="Validación de Articulos", _
           Buttons:=vbInformation
    Cnt.RollbackTrans
    trwPaquetes.SetFocus
    Exit Sub
  Case Is = 13
    MsgBox Prompt:="Las fechas de duración del plato no son validas", _
           Title:="Validación de Fechas", _
           Buttons:=vbInformation
    mskFechaIni.SetFocus
    Exit Sub
  Case Is = 6
    MsgBox Prompt:="Posiblemente no ha establecido las cantidades " & _
                   "correctamente, verifique e intente nuevamente", _
           Title:="Validación de Fechas", _
           Buttons:=vbInformation
     OSNewFlexGrid1.SetFocus
    Exit Sub
  Case Else
  MsgBox DeleteMsg(Err.Description)
End Select
Cnt.RollbackTrans
trwPaquetes.SetFocus
End Sub
'*************************************************
'FUNCIÓN PARA OBTENER LA EQUIVALENCIA DEL ARTICULO
'*************************************************
Private Function GetEquivArti(ByVal intCodArt As Long) As Long
Dim SQL As String
Dim rs As New ADODB.Recordset
SQL = "SELECT PA_EQUIV FROM PAQUETES " & _
      "WHERE PA_CODCIA='" & LK_CODCIA & "' AND " & _
      "PA_CODPA='" & Val(trwPaquetes.SelectedItem.Key) & "' AND " & _
      "PA_CODART='" & intCodArt & "'"
rs.Open SQL, Cnt, adOpenForwardOnly, adLockOptimistic, adCmdText
GetEquivArti = rs!PA_EQUIV
Set rs = Nothing
End Function
'***************************************************
'FUNCIÓN PARA OBTENER EL SIGUIENTE NUMERO DE FACTURA
'***************************************************
Private Function GetNextNumberFac() As Integer
SQ_OPER = 3
PU_TIPMOV = 70
pu_codcia = LK_CODCIA
PU_NUMSER = 1
PU_FBG = ""
LEER_FAR_LLAVE

If Not far_menor.EOF Then
   GetNextNumberFac = Val(far_menor!far_numfac) + 1
Else
 GetNextNumberFac = 1
End If

End Function
'*****************************************************
'FUNCIÓN PARA OBTENER EL SIGUIENTE NUMERO DE OPERACIÓN
'*****************************************************
Private Function GetNumberOper() As Integer
SQ_OPER = 2
PUB_FECHA = LK_FECHA_DIA
pu_codcia = LK_CODCIA
LEER_ALL_LLAVE
If all_menor.EOF = False Then
   GetNumberOper = CInt(all_menor!ALL_NUMOPER)
Else
   GetNumberOper = 0
End If
 GetNumberOper = GetNumberOper + 1
End Function
Private Sub cmdSetItem_Click()
On Error Resume Next
fraFind.Visible = True
optIP(1).Enabled = False
fraFamilias.Enabled = False
txtFindItem.Enabled = True
txtFindItem = ""
Flag_Set_Find = 0
PrmDiv = Null
PrmFam = Null
PrmLinea = Null
PrmSublinea = Null
PrmMarca = Null
PrmClase = Null
PrmDesc = Null
PrmAlterno = Null
lvwItems.ListItems.Clear
'LoadHeaders_Articulos
txtFindItem.SetFocus
End Sub

Private Sub cmdTodos_Click()
PrmDiv = Null
PrmFam = Null
PrmLinea = Null
PrmSublinea = Null
PrmMarca = Null
PrmClase = Null
PrmDesc = Null
PrmAlterno = Null
cboDivision.ListIndex = -1
cboFamilia.ListIndex = -1
cboLinea.ListIndex = -1
cboSubLinea.ListIndex = -1
cboMarca.ListIndex = -1
cboClase.ListIndex = -1
If OSNewFlexGrid1.Rows > 1 Then
 OSNewFlexGrid1.Clear
End If

Requery
trwPaquetes.SetFocus
End Sub

Private Sub Command1_Click()
Me.Hide
Unload Me
End Sub
'************************
'PROCEDIMIENTO PRINCIPAL
'************************
Private Sub Form_Load()
Dim SQL As String
Dim strCnn As String
On Error GoTo ErrorCnn
  SQL = "Exec sp_List_Pack '" & LK_CODCIA & "','P'"
         
         strCnn = "Provider=MSDataShape;Data Provider=MSDASQL;Driver={SQL Server};"
         strCnn = strCnn + "server=" & Trim(CONST_SERVER) & ""
         strCnn = strCnn + ";uid=" & Trim(CONST_UID) & ""
         strCnn = strCnn + ";pwd=" & Trim(CONST_PWD) & ""
         strCnn = strCnn + ";database=BDATOS"
         
  Cnt.ConnectionTimeout = 15
  Cnt.Open strCnn
  
  With trwPaquetes
   .FullRowSelect = True
   .ImageList = imgPaquetes
   .LabelEdit = tvwManual
   .Style = tvwTreelinesPictureText
   .LineStyle = tvwTreeLines
  
  End With
  'ESTABLECER COMANDOS
  set_Cmd
  
  PrmCodCia = LK_CODCIA
  PrmTipReg = 122
  
  Load_Familias cboDivision
  
  Requery
  
  LoadHeaderItem
Exit Sub
ErrorCnn:
 MsgBox "Imposible mostrar este formulario", vbInformation, "Error General de Red"
 Unload frmPaquete
End Sub
'************************************************************
'PROCEDIMIENTO PARA RELLENAR EL ARBOL DE ARTICULOS O PAQUETES
'************************************************************
Sub Requery()
Dim Clave As String
On Error GoTo ErrorHandle
If optIP(0).Value Then
 PrmPack = "P"
ElseIf optIP(1).Value And optIP(1).Enabled = True Then
 PrmPack = "M"
End If
'EJECUTAR COMANDO
Set rsP = CmdIp.Execute
With trwPaquetes
  .Nodes.Clear
  Set oNodo = .Nodes.Add(Text:=IIf(optIP(0), "Platos", "Insumos"), Image:="raiz")
  While Not rsP.EOF
    Clave = rsP!ART_KEY & "A"
    Set oNodo = .Nodes.Add(relative:=1, relationship:=tvwChild, _
    Key:=Clave, Text:=rsP!ART_ALTERNO & " " & rsP!ART_NOMBRE, Image:=IIf(optIP(0), "paquete", "item"))
    rsP.MoveNext
  Wend
  rsP.Close
End With
oNodo.EnsureVisible
Exit Sub
ErrorHandle:
 MsgBox Err.Description
End Sub
Private Sub Form_Unload(Cancel As Integer)
'rsP.Close
Flag_Set_Find = 0
Set rsP = Nothing
'cnT.Close
Set Cnt = Nothing
Set CmdFamilias = Nothing
Set CmdIp = Nothing
Set PrmCodArt = Nothing
Set PrmTipReg = Nothing
Set PrmCodArt = Nothing
Set PrmCodCia = Nothing
Set PrmPack = Nothing
Set PrmDiv = Nothing
Set PrmFam = Nothing
Set PrmLinea = Nothing
Set PrmSublinea = Nothing
Set PrmMarca = Nothing
Set PrmClase = Nothing

Set PrmDesc = Nothing
Set PrmAlterno = Nothing
Set rs = Nothing
End Sub
'************************
'ESTABLECER COMANDOS
'************************
Sub set_Cmd()
   With CmdFamilias
    .ActiveConnection = Cnt
    .CommandType = adCmdStoredProc
    .CommandText = "sp_list_familias"
    PrmCodCia.Type = adChar: PrmCodCia.Size = 2: PrmCodCia.Direction = adParamInput
    PrmTipReg.Type = adInteger: PrmTipReg.Direction = adParamInput
    PrmCodArt.Type = adInteger:  PrmCodArt.Direction = adParamInput
   .Parameters.Append PrmCodCia
   .Parameters.Append PrmTipReg
   .Parameters.Append PrmCodArt
  End With
  With CmdIp
   .ActiveConnection = Cnt
   .CommandType = adCmdStoredProc
   .CommandText = "sp_list_pack"
   PrmPack.Type = adChar: PrmPack.Size = 1: PrmPack.Direction = adParamInput
   PrmDiv.Type = adChar: PrmDiv.Size = 4: PrmDiv.Direction = adParamInput
   PrmFam.Type = adChar: PrmFam.Size = 4: PrmFam.Direction = adParamInput
   PrmLinea.Type = adChar: PrmLinea.Size = 4: PrmLinea.Direction = adParamInput
   PrmSublinea.Type = adChar: PrmSublinea.Size = 4: PrmSublinea.Direction = adParamInput
   PrmMarca.Type = adChar: PrmMarca.Size = 4: PrmMarca.Direction = adParamInput
   PrmClase.Type = adChar: PrmClase.Size = 4: PrmClase.Direction = adParamInput
   PrmDesc.Type = adVarChar: PrmDesc.Size = 40: PrmDesc.Direction = adParamInput
   PrmAlterno.Type = adVarChar: PrmAlterno.Size = 10: PrmAlterno.Direction = adParamInput
   .Parameters.Append PrmCodCia
   .Parameters.Append PrmPack
   .Parameters.Append PrmDiv
   .Parameters.Append PrmFam
   .Parameters.Append PrmLinea
   .Parameters.Append PrmSublinea
   .Parameters.Append PrmMarca
   .Parameters.Append PrmClase
   .Parameters.Append PrmDesc
   .Parameters.Append PrmAlterno
  End With
End Sub
'*********************************************************
'PROCEDIMIENTO PARA LLENAR LAS FAMILIAS Y SUB FAMILIAS ETC
'*********************************************************
Sub Load_Familias(CboControl As ComboBox)
On Error GoTo ErrorHandle
  'EJECUTAR EL COMANDO
  Set rs = CmdFamilias.Execute
  CboControl.Clear
  While Not rs.EOF
   With CboControl
     .AddItem StrConv(Trim(rs!tab_NOMLARGO), vbProperCase)
     .ItemData(.NewIndex) = rs!TAB_NUMTAB
    rs.MoveNext
   End With
  Wend
  Exit Sub
ErrorHandle:
 Set rs = Nothing
 MsgBox Err.Description
End Sub
Private Sub lvwItems_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
Dim sql1 As String, Sql2 As String
Dim strKey As String
Select Case KeyCode
 Case Is = 27
  fraFind.Visible = False
  trwPaquetes.SetFocus
  optIP(1).Enabled = True
  fraFamilias.Enabled = True
  Exit Sub
 Case Is = 13
  If Flag_Set_Find = 10 Then
    Flag_Set_Find = 0
    fraFind.Visible = False
    lvwItems.CheckBoxes = True
    If lvwItems.ListItems.count > 0 Then
    PrmDesc = Trim(lvwItems.SelectedItem.Text)
    End If
    Requery
    trwPaquetes.SetFocus
    Exit Sub
  End If
  sql1 = "INSERT INTO PAQUETES(PA_CODCIA,PA_CODPA,PA_CODART,ALTERNO) VALUES ("
  On Error GoTo ErrorHandle
  Cnt.BeginTrans
  For i = 1 To lvwItems.ListItems.count
   If lvwItems.ListItems(i).Checked = True Then
      strKey = Val(lvwItems.ListItems(i).Key)
      Sql2 = "'" & LK_CODCIA & "','" & _
      Val(trwPaquetes.SelectedItem.Key) & "','" & _
      strKey & "','" & Trim(lvwItems.ListItems(i).ListSubItems(1).Text) & "')"
      Cnt.Execute sql1 + Sql2
   End If
  Next
  Cnt.CommitTrans
  fraFind.Visible = False
  Call ShowItems(Val(trwPaquetes.SelectedItem.Key), Trim(LK_CODCIA))
  optIP(1).Enabled = True
  PrmDesc = Null
  PrmAlterno = Null
  fraFamilias.Enabled = True
  trwPaquetes.SetFocus
End Select
Exit Sub
ErrorHandle:
  Cnt.RollbackTrans
  fraFind.Visible = False
  Select Case Err.Number
   Case Is = -2147217900
    MsgBox Prompt:="Imposible guardar los datos, posiblemente ya esten registrados", _
           Buttons:=vbInformation, _
           Title:="Establecer Items"
    optIP(1).Enabled = True
    PrmDesc = Null
    PrmAlterno = Null
    fraFamilias.Enabled = True
    Exit Sub
   Case Else
    MsgBox Err.Description
    optIP(1).Enabled = True
    PrmDesc = Null
   PrmAlterno = Null
   fraFamilias.Enabled = True
  End Select
End Sub
Private Sub lvwItems_LostFocus()
fraFamilias.Enabled = True
End Sub
Private Sub optIP_Click(Index As Integer)
Select Case Index
  Case Is = 0
   cmdSetItem.Enabled = True
   PrmPack = "P"
   PrmDesc = Null
   PrmAlterno = Null
   If OSNewFlexGrid1.Rows > 1 Then
    OSNewFlexGrid1.Clear
   End If
   Requery
   trwPaquetes.SetFocus
  Case Is = 1
   cmdSetItem.Enabled = False
   lvwItems.ListItems.Clear
   fraFind.Visible = False
   PrmPack = "M"
   'ojo
   If OSNewFlexGrid1.Rows > 1 Then
    OSNewFlexGrid1.Clear
   End If
   If Not cboDivision = "" Then
    PrmDiv = Trim(str(Val(cboDivision.ItemData(cboDivision.ListIndex))))
   End If
   Requery
   trwPaquetes.SetFocus
End Select
End Sub
Sub LoadHeaders_Articulos()
With lvwItems
 .ColumnHeaders.Clear
 .ColumnHeaders.Add , "des", "Descripción", 3500, lvwColumnLeft
 .ColumnHeaders.Add , "cod", "Codigo", 1000, lvwColumnLeft
 .ColumnHeaders.Add , "can", "Stock", 1000, lvwColumnRight
 .LabelEdit = lvwManual
 .FullRowSelect = True
 .View = lvwReport
 .Font.Size = 8
 .Font = "Verdana"
 .CheckBoxes = True
End With
LoadArticulos
End Sub
'**************************************************
'PROCEDIMIENTO PARA PRESENTAR LA LISTA DE ARTICULOS
'**************************************************
Sub LoadArticulos()
Dim Clave As String
On Error GoTo ErrorHandle
  If optIP(0) And Flag_Set_Find = 0 Then
    PrmPack = "M"
  ElseIf optIP(0) And Flag_Set_Find = 10 Then
   PrmPack = "P"
   lvwItems.CheckBoxes = False
  ElseIf optIP(1) And Flag_Set_Find = 10 Then
   PrmPack = "M"
   lvwItems.CheckBoxes = False
  Else
    MsgBox Prompt:="Imposible mostrar la relación de items" & _
           vbCrLf & "!..seleccione mostrar platos..!", _
           Title:="Listado", _
           Buttons:=vbInformation
    fraFind.Visible = False
    fraFamilias.Enabled = True
    optIP(1).Enabled = True
    Exit Sub
  End If
  Set rsP = CmdIp.Execute
  rsP.Requery
  With lvwItems
    .ListItems.Clear
    rsP.MoveFirst
    While Not rsP.EOF
      Clave = rsP!ART_KEY & "A"
      Set oItem = .ListItems.Add(Key:=Clave, Text:=StrConv(rsP!ART_NOMBRE, vbProperCase))
      oItem.ListSubItems.Add Text:=rsP!ART_ALTERNO
      oItem.ListSubItems.Add Text:=rsP!ARM_STOCK
      If rsP!ARM_STOCK > 0 Then
       oItem.ListSubItems(2).ForeColor = RGB(0, 0, 255)
      Else
       oItem.ListSubItems(2).ForeColor = RGB(255, 0, 0)
      End If
      rsP.MoveNext
    Wend
    If .ListItems.count = 0 Then
     oItem = .ListItems.Add(Text:="No se encontro ningún Item......")
    End If
    rsP.Close
  End With
Exit Sub
ErrorHandle:
 Select Case Err.Number
 Case Is = 3021
  With lvwItems
  .CheckBoxes = False
  Set oItem = lvwItems.ListItems.Add(Text:="No se encontro ningún Item......")
  End With
 Case Else
  MsgBox Err.Description
 End Select
End Sub
Private Sub OSNewFlexGrid1_ValidateField(Cancel As Boolean, introw As Integer, intColumn As Integer)
Dim SQL As String
Select Case intColumn
 Case Is = 4
    On Error GoTo ErrorCheck
     Cnt.Execute "exec sp_check_stock '" & LK_CODCIA & "','" & _
        OSNewFlexGrid1.TextMatrix(CLng(introw), 0) & "','" & _
         Val(OSNewFlexGrid1.TEXTO) & "','" & _
         Val(trwPaquetes.SelectedItem.Key) & "'", , adExecuteNoRecords
    Cancel = False
  Exit Sub
End Select
Exit Sub
ErrorCheck:
 MsgBox DeleteMsg(Err.Description)
 Cancel = True
End Sub
Function DeleteMsg(ByVal Text As String) As String
Dim str As String * 1
Dim RES As Long
Dim New_Text As String
str = "]"
RES = InStr(1, Text, str, vbTextCompare)
New_Text = Text
If Trim(Text) = "" Or IsNull(Text) Or IsEmpty(Text) Then
 DeleteMsg = Text
 Exit Function
End If
While RES <> 0
  New_Text = Mid(New_Text, RES + 1, Len(New_Text))
  RES = InStr(1, New_Text, str, vbTextCompare)
Wend
DeleteMsg = Trim(New_Text)
End Function
Private Sub trwPaquetes_GotFocus()
OSMessage1.Text = "Listado de Platos / Presione <ALT + A>,<INSERT> ó hacer click en el boton <Agregar Insumo>  para agregar insumos al plato"
fraFind.Visible = False
fraFamilias.Enabled = True
optIP(1).Enabled = True
End Sub
Private Sub trwPaquetes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And OSNewFlexGrid1.Rows > 1 Then
 OSNewFlexGrid1.SetFocus
End If
End Sub
Private Sub trwPaquetes_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 And cmdSetItem.Enabled = True Then
 cmdSetItem = True
ElseIf KeyCode = 45 Then
 MsgBox "Imposible agregar insumos al plato", vbInformation, "Creación de Platos"
 trwPaquetes.SetFocus
 Exit Sub
End If
End Sub
Private Sub trwPaquetes_LostFocus()
OSMessage1.Text = ""
End Sub
Private Sub trwPaquetes_NodeClick(ByVal Node As MSComctlLib.Node)
Call ShowDataPack(Val(Node.Key))
Call ShowItems(Val(Node.Key), Trim(LK_CODCIA))
Debug.Print StockPack
End Sub
Private Sub txtFind_Change()
Flag_Set_Find = 10
fraFind.Visible = True
txtFindItem.Enabled = False
lvwItems.CheckBoxes = False
PrmDiv = Null
PrmFam = Null
PrmLinea = Null
PrmSublinea = Null
PrmMarca = Null
PrmClase = Null
End Sub
Private Sub txtFind_GotFocus()
OSMessage1.Text = "Ingrese la Descripción del Plato / <%> por Ayuda"
End Sub
Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
 PrmDesc = Null
 PrmAlterno = Null
 fraFind.Visible = False
 Exit Sub
End If
If KeyCode = 13 Then
 trwPaquetes.SetFocus
 Exit Sub
End If
If KeyCode = vbKeyDown Then
 If lvwItems.ListItems.count > 0 And fraFind.Visible = True Then
 lvwItems.ListItems(1).EnsureVisible
 lvwItems.ListItems(1).Selected = True
 lvwItems.SetFocus
 End If
End If

If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" And Not (KeyCode = vbKeyDown) And Not (KeyCode = vbKeyUp) Then
 If Len(txtFind) = 1 Or Len(txtFind) > 1 Then
  PrmDesc = Null
  PrmAlterno = Null
  PrmAlterno = Trim(txtFind)
  LoadHeaders_Articulos
 ElseIf Len(txtFind) = 0 Then
  With lvwItems
   .CheckBoxes = False
   .ListItems.Clear
  End With
 End If
Else
 If Len(txtFind) = 1 Or Len(txtFind) > 1 Then
  PrmAlterno = Null
  PrmDesc = Null
  PrmDesc = Trim(txtFind)
  LoadHeaders_Articulos
 ElseIf Len(txtFind) = 0 Then
  With lvwItems
   .CheckBoxes = False
   .ListItems.Clear
  End With
 End If
End If
End Sub
Private Sub txtFind_LostFocus()
OSMessage1.Text = ""
End Sub
Private Sub txtFindItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
 fraFind.Visible = False
 optIP(1).Enabled = True
 fraFamilias.Enabled = True
 PrmDesc = Null
 PrmAlterno = Null
 Exit Sub
End If
If KeyCode = vbKeyDown Then
If lvwItems.ListItems.count > 0 Then
 lvwItems.ListItems(1).EnsureVisible
 lvwItems.ListItems(1).Selected = True
End If
 lvwItems.SetFocus
End If
End Sub
'**********************************
'PRESENTAR ENCABEZADOS DE ARTICULOS
'**********************************
Sub LoadHeaderItem()
With OSNewFlexGrid1
 .FormatString = "Key|<Codigo|Descripcion|Equiv.|Cantidad|Unidad de Med.|Stock Act.|Costo de Repo.|Valor de Venta|Utilidad(%)"
 .ColEdit(0) = False
 .ColWidth(0) = 1
 .ColWidth(1) = 1000
 .ColEdit(1) = False
 .ColWidth(2) = 3000
 .ColEdit(3) = True
 .ColWidth(3) = 800
 ' cambiado el tipo de dato para la celda
 ' era entero ahora va con decimales.... JC
 '.ColData(3) = osInteger
 .ColData(3) = osDecimal
 .ColControl(3) = osTextBox
 .ColSalto(3) = 1
 .ColEdit(4) = True
 .ColData(4) = osInteger
 .ColControl(4) = osTextBox
 .ColAtras(4) = 1
 .ColSalto(4) = 1
 .ColEdit(5) = True
 .ColData(5) = osInteger
 .ColControl(5) = osComboBox
 .ColAtras(5) = 1
 .ColSalto(5) = 5
 .ColEdit(6) = Not True
 .ColEdit(7) = Not True
 .ColEdit(8) = Not True
 .ColEdit(9) = Not True
End With
End Sub
Sub ShowDataPack(ByVal codpack As Long)
Dim rs As New ADODB.Recordset
Dim SQL As String
Dim sngIGV  As Single
sngIGV = (LK_IGV / 100) + 1
SQL = "select a.art_alterno,p.pre_pre11,ar.arm_stock,ar.arm_cospro " & _
      "from arti a " & _
      "inner join precios p " & _
      "on a.art_codcia=p.pre_codcia and a.art_key=p.pre_codart " & _
      "inner join articulo ar " & _
      "on a.art_codcia=ar.arm_codcia and a.art_key=ar.arm_codart " & _
      "where a.art_key='" & codpack & "' and a.art_codcia='" & LK_CODCIA & "'"
On Error GoTo ErrorHandle
rs.Open SQL, Cnt, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not rs.EOF = True Then
 DoEvents
 StockPack = rs!ARM_STOCK
 lblPaquetes.Text = "Stock Actual de Platos: " & rs!ARM_STOCK
 If CDbl("0" & rs!arm_cospro) > 0 Then
   txtUtilidadPa = Format((((rs!pre_pre11 / sngIGV) / rs!arm_cospro) - 1) * 100, "##,##0.0000")
   txtCostoPa = Format(rs!arm_cospro, "##0.000")
   txtValorPa = Format((rs!pre_pre11 / sngIGV), "##0.000")
 Else
  txtCostoPa = "0.00"
  txtValorPa = "0.00"
  txtUtilidadPa = "0.00"
 End If
End If
Set rs = Nothing
Exit Sub
ErrorHandle:
  Set rs = Nothing
  MsgBox Err.Description
End Sub
'***************************************
'PROCEDIMIENTO PARA LLENAR LOS ARTICULOS
'DE LOS PAQUETES
'***************************************
Sub ShowItems(ByVal codpack As Long, ByVal CodCia As String)
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim sngIGV As Single
Dim cont As Long
If optIP(0) Then
  SQL = " " & _
  "Select p.pa_codcia,P.PA_EQUIV,P.PA_FLAG_ANULADO,p.pa_fecha_ini,p.pa_fecha_fin,ar.arm_stock,p.pa_prom,p.pa_codpa,a.art_key,p.pa_codart, " & _
  "p.alterno as art_alterno,p.pa_cantidad,p.pa_unidad,a.art_nombre , pr.pre_costo, pr.pre_pre11 " & _
  "from paquetes p " & _
  "inner join arti a " & _
  "on p.pa_codcia=a.art_codcia and p.pa_codart=a.art_key " & _
  "inner join precios pr " & _
  "on p.pa_codcia=pr.pre_codcia and p.pa_codart=pr.pre_codart and p.pa_unidad=pr.pre_unidad " & _
  "inner join articulo ar " & _
  "on p.pa_codcia=ar.arm_codcia and p.pa_codart=ar.arm_codart " & _
  "where pa_codpa='" & codpack & "' and pa_codcia='" & CodCia & "'"
Else
    SQL = "Select art_codcia,art_key,art_alterno, " & _
    "art_nombre , art_costo, art_cospro " & _
    "From arti " & _
    "where art_key='" & codpack & "' and art_codcia='" & CodCia & "'"
End If
On Error GoTo ErrorHandle
  rs.CursorLocation = adUseClient
  rs.Open SQL, Cnt, adOpenForwardOnly, adLockOptimistic, adCmdText
  If rs.RecordCount = 0 Then
   cmdStock.Enabled = False
   lblTitleDel.Visible = False '---JC
  Else
    cmdStock.Enabled = True '---JC
    lblTitleDel.Visible = True '---JC
  End If
  OSNewFlexGrid1.Rows = rs.RecordCount + 1
  txtTotal1 = "0.00"
  txtTotal2 = "0.00"
  txtTotal3 = "0.00"
  'Habilitar botones
  cmdActualizar.Enabled = Not False
  cmdSetItem.Enabled = Not False
  sngIGV = (LK_IGV / 100) + 1
  While Not rs.EOF
   cont = cont + 1
   With OSNewFlexGrid1
    .TextMatrix(cont, 0) = CStr(rs!ART_KEY)
    .TextMatrix(cont, 1) = CStr(rs!ART_ALTERNO)
    .TextMatrix(cont, 2) = CStr(rs!ART_NOMBRE)
    If optIP(0) Then
     If Trim(rs!PA_FLAG_ANULADO) = "S" Then
        vActu = 1
      'Deshabilitar botones
      'cmdStock.Enabled = True
      'cmdActualizar.Enabled = False
      'cmdSetItem.Enabled = False
      '.ColEdit(3) = False
      '.ColEdit(4) = False
      '.ColEdit(5) = False
      'mskFechaIni.Enabled = False
      'mskFechaFin.Enabled = False
      'txtTotal2.Enabled = False
      'txtTotal3.Enabled = False
     Else
        vActu = 0
      'mskFechaIni.Enabled = Not False
      'mskFechaFin.Enabled = Not False
      'txtTotal2.Enabled = Not False
      'txtTotal3.Enabled = Not False
      'cmdStock.Enabled = False
      '.ColEdit(3) = Not False
      '.ColEdit(4) = Not False
      '.ColEdit(5) = Not False
      'cmdActualizar.Enabled = Not False
      'cmdSetItem.Enabled = Not False
     End If
     mskFechaIni.Text = IIf(IsNull(rs!pa_fecha_ini), "__/__/____", rs!pa_fecha_ini)
     mskFechaFin.Text = IIf(IsNull(rs!pa_fecha_fin), "__/__/____", rs!pa_fecha_fin)
     .TextMatrix(cont, 3) = CStr(IIf(IsNull(rs!PA_PROM), 0, rs!PA_PROM))
     .TextMatrix(cont, 4) = CStr(IIf(IsNull(rs!PA_CANTIDAD), 0, StockPack * rs!PA_PROM))
     .TextMatrix(cont, 5) = CStr(IIf(IsNull(rs!pa_unidad), "", rs!pa_unidad))
     .TextMatrix(cont, 6) = Format(IIf(IsNull(rs!ARM_STOCK), 0, rs!ARM_STOCK) / rs!PA_EQUIV, "##,##0.00")
     .TextMatrix(cont, 7) = Format(CStr(IIf(IsNull(rs!PRE_COSTO), 0, rs!PRE_COSTO)), "##,##0.0000")
     .TextMatrix(cont, 8) = Format(CStr(IIf(IsNull(rs!pre_pre11 / sngIGV), 0, rs!pre_pre11 / sngIGV)), "##,##0.0000")
     If rs!PRE_COSTO > 0 Then
     .TextMatrix(cont, 9) = Format((((rs!pre_pre11 / sngIGV) / rs!PRE_COSTO) - 1) * 100, "##,##0.0000")
     Else
      .TextMatrix(cont, 9) = "0.000"
     End If
     txtTotal1 = Format(CDbl("0" & txtTotal1) + CDbl("0" & .TextMatrix(cont, 7) * CInt(rs!PA_PROM)), "##,##0.0000")
     txtTotal2 = Format(CDbl("0" & txtTotal2) + CDbl("0" & .TextMatrix(cont, 8) * CInt(rs!PA_PROM)), "##,##0.0000")
    End If
   End With
   rs.MoveNext
  Wend
  If CDbl("0" & txtTotal1) > 0 Then
   txtTotal3 = Format((("0" & CDbl(txtTotal2) / CDbl("0" & txtTotal1)) - 1) * 100, "##,##0.0000")
  Else
   txtTotal3 = "0.00"
  End If
   Set rs = Nothing
  Exit Sub
ErrorHandle:
Set rs = Nothing
 Select Case Err.Number
  Case Is = 11
   MsgBox "Productos nno costeados"
 Case Else
  MsgBox Err.Description
'   Resume Next
 End Select
End Sub
Private Sub txtFindItem_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
 lvwItems.ListItems(1).EnsureVisible
 lvwItems.ListItems(1).Selected = True
 lvwItems.SetFocus
 Exit Sub
End If
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" And Not (KeyCode = vbKeyDown) And Not (KeyCode = vbKeyUp) Then
 If Len(txtFindItem) = 1 Or Len(txtFindItem) > 1 Then
  PrmDesc = Null
  PrmAlterno = Null
  PrmAlterno = Trim(txtFindItem)
  LoadHeaders_Articulos
 ElseIf Len(txtFindItem) = 0 Then
  With lvwItems
   .CheckBoxes = False
   .ListItems.Clear
  End With
 End If
Else
 If Len(txtFindItem) = 1 Or Len(txtFindItem) > 1 Then
  PrmAlterno = Null
  PrmDesc = Null
  PrmDesc = Trim(txtFindItem)
  LoadHeaders_Articulos
 ElseIf Len(txtFindItem) = 0 Then
  With lvwItems
   .CheckBoxes = False
   .ListItems.Clear
  End With
 End If
End If
End Sub
Private Sub txtFindItem_LostFocus()
fraFamilias.Enabled = True
End Sub
Private Sub txtTotal2_GotFocus()
OSMessage1.Text = "Ingrese el valor de venta del plato (Sin IGV), Presione <INTRO> para visualizar los cambios"
End Sub
Private Sub txtTotal2_LostFocus()
OSMessage1.Text = ""
End Sub
Private Sub txtTotal2_ShowData(Cancel As Boolean)
If Val(txtTotal1) > 0 Then
 txtTotal3 = Format(((Val(txtTotal2) / Val(txtTotal1)) - 1) * 100, "##0.0000")
End If
End Sub
Private Sub txtTotal3_GotFocus()
OSMessage1.Text = "Ingrese el Porcentaje de Utilidad , Presione <INTRO> para visualizar los cambios"
End Sub
Private Sub txtTotal3_LostFocus()
OSMessage1.Text = ""
End Sub
Private Sub txtTotal3_ShowData(Cancel As Boolean)
txtTotal2 = Format((Val(txtTotal1) * Val(txtTotal3) / 100) + Val(txtTotal1), "##0.0000")
If Val(txtTotal3) < 0 Then
 MsgBox "Utilidad negativa verifique"
 Cancel = True
End If
End Sub
