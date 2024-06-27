VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form RCRYSTAL 
   BackColor       =   &H00FAEFDA&
   Caption         =   "Reportes en Crystal Report"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9435
   Icon            =   "frmRepoCrystal.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framoneda 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Moneda :"
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   7560
      TabIndex        =   44
      Top             =   3615
      Visible         =   0   'False
      Width           =   1455
      Begin VB.ComboBox cmdmoneda 
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
         Height          =   315
         ItemData        =   "frmRepoCrystal.frx":0442
         Left            =   120
         List            =   "frmRepoCrystal.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Banco :"
      Height          =   1095
      Left            =   0
      TabIndex        =   40
      Top             =   2160
      Visible         =   0   'False
      Width           =   3135
      Begin VB.TextBox txt_key 
         DataField       =   "ART_KEY"
         DataSource      =   "Data1"
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
         MaxLength       =   8
         TabIndex        =   41
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblbanco 
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
         Height          =   435
         Left            =   120
         TabIndex        =   42
         Top             =   615
         Width           =   2925
         WordWrap        =   -1  'True
      End
   End
   Begin VB.ListBox SITUACION 
      Height          =   1185
      Left            =   4800
      Style           =   1  'Checkbox
      TabIndex        =   38
      Top             =   1320
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.ListBox TIPDOC 
      Height          =   1185
      Left            =   3600
      Style           =   1  'Checkbox
      TabIndex        =   37
      Top             =   1320
      Visible         =   0   'False
      Width           =   1020
   End
   Begin ComctlLib.ListView ListView2 
      Height          =   375
      Left            =   7080
      TabIndex        =   33
      Top             =   3720
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
   Begin VB.Frame fracodclie 
      BackColor       =   &H00FAEFDA&
      Height          =   1095
      Left            =   2535
      TabIndex        =   34
      Top             =   1755
      Visible         =   0   'False
      Width           =   3615
      Begin VB.TextBox txt_cli 
         DataField       =   "ART_KEY"
         DataSource      =   "Data1"
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
         MaxLength       =   8
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblCliente 
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
         Height          =   435
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Visible         =   0   'False
         Width           =   3285
         WordWrap        =   -1  'True
      End
   End
   Begin VB.ListBox multiven 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   32
      Top             =   2160
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Frame fraclipro 
      Height          =   855
      Left            =   120
      TabIndex        =   29
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
      Begin VB.ComboBox cmbclipro 
         Height          =   315
         ItemData        =   "frmRepoCrystal.frx":0461
         Left            =   120
         List            =   "frmRepoCrystal.frx":046B
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblclipro 
         AutoSize        =   -1  'True
         Caption         =   "Cliente / Proveedor"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1380
      End
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   375
      Left            =   6120
      TabIndex        =   28
      Top             =   3720
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
   Begin VB.Frame frafechas 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Fechas :"
      Height          =   855
      Left            =   120
      TabIndex        =   20
      Top             =   1200
      Visible         =   0   'False
      Width           =   2775
      Begin MSMask.MaskEdBox txtCampo2 
         Height          =   285
         Left            =   1440
         TabIndex        =   21
         Top             =   480
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
         TabIndex        =   22
         Top             =   480
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
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   585
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
         Left            =   1440
         TabIndex        =   23
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame frazonas 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Filtro para Clientes :"
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
      Height          =   3255
      Left            =   0
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   9015
      Begin VB.CheckBox cheestado 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Mostrar Desactivos"
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
         Left            =   5040
         TabIndex        =   46
         Top             =   2280
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ListBox zonas 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Left            =   5640
         Style           =   1  'Checkbox
         TabIndex        =   18
         Top             =   240
         Width           =   3255
      End
      Begin VB.OptionButton opzonas 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Distrito"
         Height          =   240
         Index           =   0
         Left            =   3000
         TabIndex        =   17
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton opzonas 
         Caption         =   "Provincia"
         Height          =   240
         Index           =   1
         Left            =   3000
         TabIndex        =   16
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton opzonas 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Zonas"
         Height          =   240
         Index           =   2
         Left            =   3000
         TabIndex        =   15
         Top             =   1320
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.Label lblzonas 
         AutoSize        =   -1  'True
         Caption         =   "Zonas :"
         Height          =   195
         Left            =   4080
         TabIndex        =   19
         Top             =   600
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Formulas a Incluir :"
      Height          =   855
      Left            =   4920
      TabIndex        =   12
      Top             =   0
      Width           =   4095
      Begin VB.Label lblformulas 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   3855
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraArti 
      Caption         =   "Filtro de Articulos :"
      Height          =   3255
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   9015
      Begin VB.ListBox art_marca 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   6120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   48
         Top             =   1800
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ListBox art_numero 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   6120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   47
         Top             =   240
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox i_codart2 
         Height          =   315
         Left            =   240
         MaxLength       =   8
         TabIndex        =   25
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox famix 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   3000
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.ListBox subfami 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Left            =   6120
         Style           =   1  'Checkbox
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ListBox fami 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Left            =   3000
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label i_nomarti 
         AutoSize        =   -1  'True
         Caption         =   "             "
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
         Left            =   1560
         TabIndex        =   27
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label lblarti 
         Caption         =   "Articulo :"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   26
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblarti 
         Caption         =   "Sub. Familias  :"
         Height          =   255
         Index           =   1
         Left            =   6000
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblarti 
         Caption         =   "Familias :"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Frame fra1 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Reporte :"
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4815
      Begin VB.Label lblreporte 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.CommandButton cmdCerrar 
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
      Left            =   5070
      TabIndex        =   1
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton Pantalla 
      Caption         =   "&Pantalla"
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
      Left            =   2670
      TabIndex        =   0
      Top             =   4920
      Width           =   1695
   End
   Begin ComctlLib.ProgressBar ProgBar 
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   120
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.ListView ListView3 
      Height          =   375
      Left            =   6960
      TabIndex        =   43
      Top             =   4440
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin VB.Label LBLTIPDOC 
      Caption         =   "Tipo de Documentos y Situación"
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
      Left            =   3720
      TabIndex        =   39
      Top             =   960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblproceso 
      BackStyle       =   0  'Transparent
      Caption         =   "Procesando . . ."
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
      Left            =   2400
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   2775
   End
End
Attribute VB_Name = "RCRYSTAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Wfile As String
Dim WFORM As String
Dim REP_FECHA1 As String
Dim REP_FECHA2 As String
Dim VAR_ACTIVAR As Integer
Dim WCOD_ORIGINAL As Currency
Dim loc_key As Integer
Dim loc_cp As String * 1
Dim WW_CODVEN As Integer


Private Sub cmdcerrar_Click()
Unload RCRYSTAL
End Sub

Private Sub famix_Click()
Dim wpos As Integer
Dim WFAMI2 As Integer
'If Flag_Bloq = "A" Then
' Exit Sub
'End If
If Trim(famix.Text) = "" Then
 subfami.Clear
 Exit Sub
End If
wpos = subfami.ListIndex
WFAMI2 = Val(Trim(Right(famix.Text, 6)))
LLENADO_SUBFAM WFAMI2
On Error GoTo SIGUE
subfami.ListIndex = wpos
Exit Sub
SIGUE:
Resume Next

End Sub

Private Sub Form_Load()
VAR_ACTIVAR = 0
CenterMe RCRYSTAL
Screen.MousePointer = 11
If tra_llave.EOF Then
   Screen.MousePointer = 0
   Exit Sub
End If
Screen.MousePointer = 0
Wfile = Trim(tra_llave(3))
WFORM = Trim(tra_llave(7))

If tra_llave!tra_con6 = 1 Then
  i_codart2.TabIndex = 0
  fraArti.Visible = True
  lblarti(2).Visible = True
  i_codart2.Visible = True
ElseIf tra_llave!tra_con7 = 1 Then
 PUB_CODCIA = LK_CODCIA
 If LK_EMP_PTO = "A" Then
      PUB_CODCIA = "00"
 End If
 LLENADOS fami, 122
 fami.TabIndex = 0
 lblarti(0).Visible = True
 subfami.TabIndex = 1
 fraArti.Visible = True
 fami.Visible = True
ElseIf tra_llave!tra_con9 = 1 Then
 PUB_CODCIA = LK_CODCIA
 If LK_EMP_PTO = "A" Then
      PUB_CODCIA = "00"
 End If
 LLENADOS famix, 122
 famix.TabIndex = 0
 fami.Visible = False
 lblarti(0).Visible = True
 lblarti(1).Visible = True
 subfami.TabIndex = 1
 subfami.Visible = True
 famix.Visible = True
 fraArti.Visible = True
End If
If tra_llave!tra_act12 = 1 Then
 PUB_CODCIA = LK_CODCIA
 LLENADOS art_numero, 130
 art_numero.Visible = True
 fraArti.Visible = True
End If
If tra_llave!tra_act13 = 1 Then
 PUB_CODCIA = LK_CODCIA
 LLENADOS art_marca, 132
 art_marca.Visible = True
 fraArti.Visible = True
End If


If tra_llave!TRA_CON8 = 1 Then
 PUB_CODCIA = "00"
 LLENADOS zonas, 35
 frazonas.Visible = True
 opzonas(0).Caption = BUSCA_ETIQUETA(10)
 opzonas(1).Caption = BUSCA_ETIQUETA(11)
 opzonas(2).Caption = BUSCA_ETIQUETA(12)
End If
If tra_llave!TRA_CON4 = 1 Then
  fracodclie.Caption = "Cliente "
  fracodclie.Visible = True
  txt_cli.Visible = True
  txt_cli.TabIndex = 0
  lblCliente.Visible = True
  loc_cp = "C"
End If
If tra_llave!tra_CON5 = 1 Then
  txt_cli.TabIndex = 0
  fracodclie.Caption = "Proveedor "
  fracodclie.Visible = True
  txt_cli.Visible = True
  lblCliente.Visible = True
  loc_cp = "P"
End If

If tra_llave!TRA_CON11 = 1 Then
  fraclipro.Visible = True
  cmbclipro.ListIndex = 0
End If
If tra_llave!tra_ACT5 = 1 Or tra_llave!tra_con14 = 1 Or tra_llave!tra_con1 = 1 Or tra_llave!tra_con10 = 1 Or tra_llave!tra_act8 = 1 Or tra_llave!tra_con12 = 1 Then
 frafechas.Visible = True
 lblcampo1.Caption = "Fecha de Inicial : "
 txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtCampo1.Mask = "##/##/####"
 lblcampo2.Caption = "Fecha de Final: "
 txtCampo2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtCampo2.Mask = "##/##/####"
 txtCampo1.TabIndex = 0
 txtCampo2.TabIndex = 1
End If
If tra_llave!TRA_CON3 = 1 Then
  LLENA_VENDEDORES
  multiven.Visible = True
  multiven.TabIndex = 1
End If
If tra_llave!TRA_CON2 = 1 Then
 Frame2.Visible = True
End If
If tra_llave!TRA_ACT6 = 1 Then
  framoneda.Visible = True
  cmdmoneda.ListIndex = 0
End If
If tra_llave!TRA_ACT7 = 1 Then
  cheestado.Visible = True
End If

lblformulas.Caption = ""
If tra_llave!tra_ACT1 = 1 Then lblformulas.Caption = lblformulas.Caption + "; CIA   "
If tra_llave!tra_ACT2 = 1 Then lblformulas.Caption = lblformulas.Caption + "; DIA   "
If tra_llave!tra_ACT10 = 1 Then lblformulas.Caption = lblformulas.Caption + "; FECHAS   "
If tra_llave!TRA_ACT4 = 1 Then
 LBLTIPDOC.Visible = True
 SITUACION.Visible = True
 TIPDOC.Visible = True
 PUB_CODCIA = "00"
 LLENADOS SITUACION, 133
 LLENADOS TIPDOC, 8
End If

lblreporte.Caption = Trim(tra_llave(1))
If Wfile = "SALINI.RPT" Then
  txtCampo1.Enabled = False
  txtCampo2.Enabled = False
  Pantalla.TabIndex = 0
End If
End Sub

Private Sub i_codart2_Change()
If i_codart2.Text = "" Then
 i_nomarti.Caption = ""
  VAR_ACTIVAR = 0
End If

End Sub

Private Sub i_codart2_GotFocus()
'Azul i_codart2, i_codart2
'i_codart2.text = ""
'i_nomarti.Caption = ""
End Sub
Private Sub i_codart2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And i_codart2.Text = "" Then
  loc_key = 1
  Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView1.ListItems.Count Then loc_key = ListView1.ListItems.Count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView1.ListItems.Count Then loc_key = ListView1.ListItems.Count
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
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  DoEvents
  i_codart2.SelStart = Len(i_codart2.Text)
  DoEvents
fin:

End Sub
Private Sub i_codart2_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i, car
Dim itmFound As ListItem
car = Chr(KeyAscii)
KeyAscii = Asc(UCase(car))
If KeyAscii = 27 Then
 ListView1.Visible = False
 i_codart2.Text = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
VAR_ACTIVAR = 0
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  PUB_KEY = 0
Else
 On Error GoTo mucho
 PUB_KEY = Val(i_codart2.Text)
 On Error GoTo 0
 If Len(i_codart2.Text) = 0 Then
    Exit Sub
 End If
 If IsNumeric(i_codart2.Text) = False Then
   PUB_KEY = 0
 End If
End If

If PUB_KEY <> 0 Then
    SQ_OPER = 1
    PUB_KEY = i_codart2.Text
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then
       MsgBox "Codigo NO Existe.", 48, Pub_Titulo
       Azul i_codart2, i_codart2
       GoTo fin
    End If
    WCOD_ORIGINAL = art_LLAVE!ART_KEY
    i_nomarti.Caption = Trim(art_LLAVE!ART_NOMBRE)
    ListView1.Visible = False
    Pantalla.SetFocus
    Exit Sub
Else
  If ListView1.Visible = False And VAR_ACTIVAR <> 99 And i_codart2.Text <> "" And LK_FLAG_ORIGINAL <> "A" And LK_FLAG_ALTERNO = "A" Then
IR_ALTERNO:
     SQ_OPER = 3
     pu_alterno = i_codart2.Text
     pu_codcia = LK_CODCIA
     LEER_ART_LLAVE
     If art_llave_alt.EOF Then
       MsgBox "Codigo No Existe ...", 48, Pub_Titulo
       Azul i_codart2, i_codart2
       Exit Sub
     End If
     WCOD_ORIGINAL = art_llave_alt!ART_KEY
     'i_codart2.text = Trim(art_llave_alt!ART_NOMBRE)
     i_nomarti.Caption = Trim(art_llave_alt!ART_NOMBRE)
     ListView1.Visible = False
     Pantalla.SetFocus
     Exit Sub
  Else
    If loc_key > ListView1.ListItems.Count Or loc_key = 0 Then
     Exit Sub
    End If
    valor = UCase(ListView1.ListItems.Item(loc_key).Text)
    If Trim(UCase(i_codart2.Text)) = Left(valor, Len(Trim(i_codart2.Text))) And Len(Trim(i_codart2.Text)) <> 0 Then
      If VAR_ACTIVAR = 0 And LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
        i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key))
        GoTo IR_ALTERNO
      End If
      If VAR_ACTIVAR <> 99 Then
       i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
      Else
       i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key))
      End If
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       PUB_KEY = Val(ListView1.ListItems.Item(loc_key).SubItems(1))
      Else
       PUB_KEY = i_codart2.Text
      End If
      LEER_ART_LLAVE
      VAR_ACTIVAR = 0
      If art_LLAVE.EOF Then
        MsgBox "Codigo No Existe ...", 48, Pub_Titulo
        Azul i_codart2, i_codart2
        Exit Sub
      End If
      WCOD_ORIGINAL = art_LLAVE!ART_KEY
      i_nomarti.Caption = Trim(art_LLAVE!ART_NOMBRE)
      i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
      ListView1.Visible = False
      Pantalla.SetFocus
      Exit Sub
    Else
      Exit Sub
    End If
    
  End If
End If
dale:
ListView1.Visible = False
fin:
mucho:
Exit Sub
ERROR_CODIGO:
MsgBox "Codigo NO Valido .... ", 48, Pub_Titulo
Azul i_codart2, i_codart2
  

End Sub

Private Sub i_codart2_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
If KeyCode = 13 Then Exit Sub
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  If Len(i_codart2.Text) = 0 Or i_codart2.Text = "" Then
    ListView1.Visible = False
    Exit Sub
  End If
  If i_codart2.Text = "*" And KeyCode = 106 Then
   VAR_ACTIVAR = 99
   Exit Sub
  ElseIf i_codart2.Text = "" Then
   VAR_ACTIVAR = 0
   Exit Sub
  End If
  If VAR_ACTIVAR <> 99 Then
    Exit Sub
  End If
  If Left(i_codart2.Text, 1) = "*" Then
   i_codart2.Text = Mid(i_codart2.Text, 2, Len(i_codart2.Text))
   i_codart2.SelStart = Len(i_codart2.Text)
  End If
Else
 If Len(i_codart2.Text) = 0 Or IsNumeric(i_codart2.Text) = True Then
   ListView1.Visible = False
   Exit Sub
 End If
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(i_codart2.Text) = 1 Then
    var = Asc(i_codart2.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      numarchi = 3
      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_CODCIA = '" & LK_CODCIA & "' AND ART_ALTERNO BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
    Else
      numarchi = 0
      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_CODCIA = '" & LK_CODCIA & "' AND ART_NOMBRE BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
    End If
    PROC_LISVIEW ListView1
    loc_key = 0
    If ListView1.Visible Then
     loc_key = 1
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(i_codart2.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView1.ListItems.Count Then
      ListView1.ListItems.Item(ListView1.ListItems.Count).EnsureVisible
   Else
     ListView1.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If



End Sub



Private Sub i_codart2_LostFocus()
ListView1.Visible = False
End Sub
Private Sub ListView1_DblClick()
 loc_key = ListView1.SelectedItem.Index
 i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
 i_codart2_KeyPress 13

End Sub

Private Sub ListView1_GotFocus()
If loc_key <> 0 Then
 Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
 ListView1.ListItems.Item(loc_key).Selected = True
 ListView1.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = ListView1.SelectedItem.Index
 i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
End If

End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
 Exit Sub
End If
i_codart2_KeyPress 13
End Sub

Private Sub ListView1_LostFocus()
ListView1.Visible = False
End Sub
Private Sub ListExiste_LostFocus()
If frmCLI.ListExiste.Visible = False Then
    Exit Sub
End If
End Sub

Private Sub ListView2_DblClick()
 loc_key = ListView2.SelectedItem.Index
 txt_cli.Text = Trim(ListView2.ListItems.Item(loc_key).Text) & " "
 txt_cli_KeyPress 13
End Sub

Private Sub ListView2_GotFocus()
If loc_key <> 0 Then
 Set ListView2.SelectedItem = ListView2.ListItems(loc_key)
 ListView2.ListItems.Item(loc_key).Selected = True
 ListView2.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub ListView2_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = ListView2.SelectedItem.Index
 txt_cli.Text = Trim(ListView2.ListItems.Item(loc_key).Text) & " "
End If

End Sub

Private Sub ListView2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 ListView2.Visible = False
 txt_cli.Text = ""
 txt_cli.SetFocus
 Exit Sub
End If
If KeyAscii <> 13 Then
 Exit Sub
End If
ListView2_DblClick

End Sub

Private Sub ListView2_LostFocus()
ListView2.Visible = False
End Sub



Private Sub multiven_Click()
If Pantalla.Enabled = True Then
If LK_EMP = "PAR" Then
 WW_CODVEN = Val(Left(multiven.Text, 3))
 PUB_CODCIA = "00"
 LLENADOS zonas, 35
 End If
End If
End Sub

Private Sub opzonas_Click(Index As Integer)
Dim cod As Integer
lblzonas.Caption = Trim(opzonas(Index).Caption) & " :"
If Index = 0 Then
  cod = 20
ElseIf Index = 1 Then
  cod = 30
ElseIf Index = 2 Then
  cod = 35
End If
PUB_CODCIA = "00"
LLENADOS zonas, cod
zonas.SetFocus

End Sub

Private Sub Pantalla_Click()
PRO_REPORTE
End Sub
Public Sub LLENADOS(cont As ListBox, tip As Integer)
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
       If PUB_TIPREG = 35 And LK_EMP = "PAR" Then
          If Val(tab_mayor!TAB_CODART) = WW_CODVEN Then
            cont.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!TAB_NUMTAB
          End If
       Else
           cont.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!TAB_NUMTAB
       End If
       tab_mayor.MoveNext
    Loop
End Sub
Public Sub LLENADO_SUBFAM(wfami As Integer)
    PUB_TIPREG = 123
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
      PUB_CODCIA = "00"
    End If
    PUB_CODART = wfami
    SQ_OPER = 3
    LEER_TAB_LLAVE
    subfami.ToolTipText = "TAB_TIPREG = 123"
    subfami.Clear
    Do Until tab_menor.EOF
        DoEvents
        subfami.AddItem tab_menor!tab_nomlargo & String(50, " ") & Trim(CStr(tab_menor!TAB_NUMTAB))
        tab_menor.MoveNext
    Loop
End Sub
Public Function SON_FECHAS() As Boolean
SON_FECHAS = True
If Right(RCRYSTAL.txtCampo1.Text, 2) = "__" Then
  REP_FECHA1 = Left(RCRYSTAL.txtCampo1.Text, 8)
Else
  REP_FECHA1 = Trim(RCRYSTAL.txtCampo1.Text)
End If
If Not IsDate(REP_FECHA1) Then
    MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
    Azul2 RCRYSTAL.txtCampo1, RCRYSTAL.txtCampo1
    GoTo fin
End If
If Right(RCRYSTAL.txtCampo2.Text, 2) = "__" Then
  REP_FECHA2 = Left(RCRYSTAL.txtCampo2.Text, 8)
Else
  REP_FECHA2 = Trim(RCRYSTAL.txtCampo2.Text)
End If
If Not IsDate(REP_FECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 Azul2 RCRYSTAL.txtCampo2, RCRYSTAL.txtCampo2
 GoTo fin
End If
If CDate(REP_FECHA1) > CDate(REP_FECHA2) Then
 MsgBox "Fechas Invalidadas ..", 48, Pub_Titulo
 Azul2 RCRYSTAL.txtCampo1, RCRYSTAL.txtCampo1
 GoTo fin
End If

Exit Function
fin:
SON_FECHAS = False

End Function

Public Sub PRO_REPORTE()
Dim xcuenta2 As Integer
Dim wf1, wf2, wf3, wf4, wf5, wf6, wf7, wf8, wf9, wf10
Dim DIA, MES, ano
Dim DIA1, MES1, ANO1
Dim PU_MONEDA As String * 1
Dim WFECHA_DIA As String
Dim wfiltra As String
Dim wmensa  As String
Dim CADENITA, Modo1 As String
Dim warma_arti As String
Dim wcodcia As String * 2
Dim M1, A1 As Integer
Dim M2, A2 As Integer
Dim M3, A3 As Integer
Dim M4, A4 As Integer
Dim M5, A5 As Integer
Dim M6, A6 As Integer
Dim M7, A7 As Integer
Dim M8, A8 As Integer
Dim M9, A9 As Integer
Dim M10, A10 As Integer
Dim M11, A11 As Integer
Dim M12, A12 As Integer


On Error GoTo SALE
' <<< CONSISTENCIAS >>>
If tra_llave!TRA_CON4 = "1" Or tra_llave!tra_CON5 = "1" Then
  If Trim(txt_cli.Text) = "" Then
      MsgBox "Verificar Codigo ", 48, Pub_Titulo
      Exit Sub
  End If
End If



wf1 = ""
wf2 = ""
wf3 = ""
wf4 = ""
wf5 = ""
wf6 = ""
wf7 = ""
wf8 = ""
wf9 = ""
wf10 = ""
Pantalla.Enabled = False
cmdcerrar.Enabled = False

Screen.MousePointer = 11
ProgBar.Min = 0
ProgBar.Max = 10
ProgBar.Value = 0
ProgBar.Visible = True
lblproceso.Visible = True
DoEvents
If Len(Wfile) = 0 Then
 MsgBox " Cheque los datos de Reportes , Intente nuevamente.", 48, Pub_Titulo
 Exit Sub
End If
  ProgBar.Value = 2
  Reportes.Connect = PUB_ODBC
  If tra_llave!tra_rep1 = "1" Then
     Reportes.ReportFileName = PUB_RUTA_OTRO & "PTOVTA\" & Wfile
     wcodcia = LK_CODCIA
  Else
    Reportes.ReportFileName = PUB_RUTA_OTRO & Wfile
    wcodcia = LK_CODCIA
  End If
  Reportes.WindowTitle = "Reporte :  " & Trim(tra_llave(1)) & " - Archivo:(" & Wfile & ")"
  ProgBar.Value = 4
  Reportes.Destination = crptToWindow
  Reportes.WindowLeft = 2
  Reportes.WindowTop = 70
  Reportes.WindowWidth = 635
  Reportes.WindowHeight = 390
  Reportes.Formulas(0) = ""
  Reportes.Formulas(1) = ""
  Reportes.Formulas(2) = ""
  Reportes.Formulas(3) = ""
  Reportes.Formulas(4) = ""
  Reportes.Formulas(5) = ""
  Reportes.Formulas(6) = ""
  Reportes.Formulas(7) = ""
  Reportes.Formulas(8) = ""
  Reportes.Formulas(9) = ""
  Reportes.Formulas(10) = ""
  ProgBar.Value = 6
  pub_cadena = ""
  wmensa = ""
  If tra_llave!TRA_CON2 = 1 Then
    If pub_cadena = "" Then
       pub_cadena = "{CCMAEST.CCM_CODBAN} = " & Trim(txt_key.Text)
    Else
        pub_cadena = pub_cadena + " AND " + "{CCMAEST.CCM_CODBAN} = " & Trim(txt_key.Text)
    End If
    If pub_cadena = "" Then
          pub_cadena = pub_cadena + " {CCMAEST.CCM_CODCIA} = " + "'" & LK_CODCIA & "' "
    Else
          pub_cadena = pub_cadena + " AND " + " {CCMAEST.CCM_CODCIA} = " + "'" & LK_CODCIA & "' "
    End If
  End If
  If tra_llave!TRA_S1 = 1 Then
    If pub_cadena = "" Then
       pub_cadena = "{COMAEST.COM_CODCIA} = '" & LK_CODCIA & "'"
    Else
        pub_cadena = pub_cadena + " AND " + "{COMAEST.COM_CODCIA} = '" & LK_CODCIA & "'"
    End If
  End If

  If tra_llave!TRA_CON4 = 1 Or tra_llave!tra_CON5 = 1 Then
    If pub_cadena = "" Then
       pub_cadena = "{CLIENTES.CLI_CODCLIE} = " & Trim(txt_cli.Text)
    Else
        pub_cadena = pub_cadena + " AND " + "{CLIENTES.CLI_CODCLIE} = " & Trim(txt_cli.Text)
    End If
    If pub_cadena = "" Then
          pub_cadena = pub_cadena + " {CLIENTES.CLI_CP} = '" + loc_cp + "' AND {CLIENTES.CLI_CODCIA} = " + "'" & LK_CODCIA & "' "
    Else
          pub_cadena = pub_cadena + " AND " + " {CLIENTES.CLI_CP} = '" + loc_cp + "' AND {CLIENTES.CLI_CODCIA} = " + "'" & LK_CODCIA & "' "
    End If
  End If
  If tra_llave!tra_con6 = 1 Then  ' X articulo
   If LK_EMP_PTO = "A" Then
     warma_arti = " {ARTICULO.ARM_CODCIA} = "
   Else
     warma_arti = " {ARTI.ART_CODCIA} = "
   End If
   If pub_cadena = "" Then
       pub_cadena = "{ARTI.ART_KEY} = " & Str(WCOD_ORIGINAL)
   Else
       pub_cadena = pub_cadena + " AND " + "{ARTI.ART_KEY} = " & Str(WCOD_ORIGINAL)
   End If
   If tra_llave!tra_ACT11 = 0 Then
    If pub_cadena = "" Then
          pub_cadena = pub_cadena + warma_arti + "'" & LK_CODCIA & "' "
    Else
          pub_cadena = pub_cadena + " AND " + warma_arti + "'" & LK_CODCIA & "' "
    End If
   End If
  ElseIf tra_llave!tra_con7 = 1 Then ' x FAMI
      If LK_EMP_PTO = "A" Then
        warma_arti = " {ARTICULO.ARM_CODCIA} = "
      Else
        warma_arti = " {ARTI.ART_CODCIA} = "
      End If
      GoSub ARMA_FAMI
      If pub_cadena = "" Then
         pub_cadena = pub_cadena + CADENITA
      Else
         pub_cadena = pub_cadena + " AND " + CADENITA
      End If
      If tra_llave!tra_ACT11 = 0 Then
       If pub_cadena = "" Then
          pub_cadena = pub_cadena + warma_arti + "'" & LK_CODCIA & "' "
       Else
          pub_cadena = pub_cadena + " AND " + warma_arti & "'" & LK_CODCIA & "' "
       End If
      Else
        xcuenta2 = 1
        CADENITA = ""
        For fila = 1 To 30
          pu_codcia = Mid(Trim(LK_ART_CIAS), xcuenta2, 2)
          If Trim(pu_codcia) = "" Then Exit For
          CADENITA = CADENITA + " {ARTI.ART_CODCIA} = '" & pu_codcia & "' OR "
          xcuenta2 = xcuenta2 + 2
        Next fila
        If CADENITA <> "" Then
          CADENITA = "(" & Mid(CADENITA, 1, Len(CADENITA) - 4) & ")"
          pub_cadena = pub_cadena + " AND  " & CADENITA
        End If
      End If
      wmensa = wmensa + "Fam.: " + wfiltra
  ElseIf tra_llave!TRA_CON8 = 1 Then
     GoSub ARMA_ZONA:
     If pub_cadena = "" Then
         pub_cadena = pub_cadena + CADENITA
     Else
        If CADENITA <> "" Then
         pub_cadena = pub_cadena + " AND " + CADENITA
        End If
     End If
     If pub_cadena = "" Then
         pub_cadena = pub_cadena + " {CLIENTES.CLI_CODCIA} = '" & LK_CODCIA & "' "
      Else
         pub_cadena = pub_cadena + " AND  {CLIENTES.CLI_CODCIA} = '" & LK_CODCIA & "'  "
      End If
     wmensa = wmensa + Trim(lblzonas.Caption) + wfiltra
  ElseIf tra_llave!tra_con9 = 1 Then ' x FAMI SUB FAMI
      If pub_cadena = "" Then
         pub_cadena = "{ARTI.ART_FAMILIA} in [" & Str(Val(Right(famix.Text, 6))) & "]"
      Else
         pub_cadena = pub_cadena + " AND " + "{ARTI.ART_FAMILIA} in [" & Str(Val(Right(fami.Text, 6))) & "]"
      End If
      wmensa = wmensa + "Fam.: " + Left(famix.Text, 8)
      GoSub ARMA_SUBFAMI:
      If pub_cadena = "" Then
         pub_cadena = pub_cadena + CADENITA
      Else
        If CADENITA <> "" Then
         pub_cadena = pub_cadena + " AND " + CADENITA
        End If
      End If
     If tra_llave!tra_ACT11 = 0 Then
       If LK_EMP_PTO = "A" Then
          warma_arti = " {ARTICULO.ARM_CODCIA} = "
       Else
          warma_arti = " {ARTI.ART_CODCIA} = "
       End If
       If pub_cadena = "" Then
         pub_cadena = pub_cadena + warma_arti + "'" & LK_CODCIA & "' "
       Else
          pub_cadena = pub_cadena + " AND " + warma_arti + "'" & LK_CODCIA & "' "
       End If
    
     End If
      wmensa = wmensa + "Sub.Fam.: " + wfiltra
  End If
  If tra_llave!tra_act12 = 1 Then
      If LK_EMP_PTO = "A" Then
        warma_arti = " {ARTICULO.ARM_CODCIA} = "
      Else
        warma_arti = " {ARTI.ART_CODCIA} = "
      End If
      GoSub ARMA_NUMERO
      If CADENITA <> "" Then
        If pub_cadena <> "" Then
           pub_cadena = pub_cadena + " AND " + CADENITA
        Else
           pub_cadena = pub_cadena + CADENITA
        End If
      End If
      If tra_llave!tra_ACT11 = 0 Then
       If pub_cadena = "" Then
          pub_cadena = pub_cadena + warma_arti + "'" & LK_CODCIA & "' "
       Else
          pub_cadena = pub_cadena + " AND " + warma_arti & "'" & LK_CODCIA & "' "
       End If
      End If
  End If
      
  If tra_llave!tra_ACT5 = 1 Then ' x FECHAS X CHEQUES
    If Not SON_FECHAS Then
    GoTo SALE
    End If
    DIA = Day(REP_FECHA1)
    MES = Month(REP_FECHA1)
    ano = Year(REP_FECHA1)
    DIA1 = Day(REP_FECHA2)
    MES1 = Month(REP_FECHA2)
    ANO1 = Year(REP_FECHA2)
    CADENITA = "{CHEQUES.CHE_FECHA} >= Date ( " & ano & "," & MES & "," & DIA & ") AND {CHEQUES.CHE_FECHA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
    If pub_cadena = "" Then
       pub_cadena = pub_cadena + CADENITA
    Else
      pub_cadena = pub_cadena + " AND " + CADENITA
    End If
    If pub_cadena = "" Then
         pub_cadena = pub_cadena + " {CHEQUES.CHE_CODCIA} = '" & LK_CODCIA & "' "
    Else
         pub_cadena = pub_cadena + " AND  {CHEQUES.CHE_CODCIA} = '" & LK_CODCIA & "' "
    End If
  End If
  If tra_llave!tra_con1 = 1 Then ' x FECHAS X FACART
    If Not SON_FECHAS Then
    GoTo SALE
    End If
    DIA = Day(REP_FECHA1)
    MES = Month(REP_FECHA1)
    ano = Year(REP_FECHA1)
    DIA1 = Day(REP_FECHA2)
    MES1 = Month(REP_FECHA2)
    ANO1 = Year(REP_FECHA2)
    If tra_llave!tra_ACT9 = 1 Then ' x FECHAS X FACART
      pub_mensaje = "Imprimir según Usuario... ?"
      Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
      If Pub_Respuesta = vbYes Then
         CADENITA = "{FACART.FAR_CODUSU}= '" & LK_CODUSU & "' AND {FACART.FAR_FECHA} >= Date ( " & ano & "," & MES & "," & DIA & ") AND {FACART.FAR_FECHA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
      Else
         If LK_EMP = "3AA" Or LK_EMP = "CAM" Or LK_EMP = "PIU" Then
            CADENITA = "{FACART.FAR_FECHA_COMPRA} >= Date ( " & ano & "," & MES & "," & DIA & ") AND {FACART.FAR_FECHA_COMPRA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
         Else
            CADENITA = "{FACART.FAR_FECHA} >= Date ( " & ano & "," & MES & "," & DIA & ") AND {FACART.FAR_FECHA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
         End If
      End If
    Else
      If LK_EMP = "3AA" Or LK_EMP = "HER" Or LK_EMP = "CAM" Or LK_EMP = "PIU" Then
          CADENITA = "{FACART.FAR_FECHA_COMPRA} >= Date ( " & ano & "," & MES & "," & DIA & ") AND {FACART.FAR_FECHA_COMPRA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
      Else
          CADENITA = "{FACART.FAR_FECHA} >= Date ( " & ano & "," & MES & "," & DIA & ") AND {FACART.FAR_FECHA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
      End If
    End If
    If pub_cadena = "" Then
       pub_cadena = pub_cadena + CADENITA
    Else
      pub_cadena = pub_cadena + " AND " + CADENITA
    End If
    If tra_llave!tra_ACT11 = 1 Then
        xcuenta2 = 1
        CADENITA = ""
        For fila = 1 To 30
          pu_codcia = Mid(Trim(LK_ART_CIAS), xcuenta2, 2)
          If Trim(pu_codcia) = "" Then Exit For
          CADENITA = CADENITA + " {FACART.FAR_CODCIA} = '" & pu_codcia & "' OR "
          xcuenta2 = xcuenta2 + 2
        Next fila
        If CADENITA <> "" Then
          CADENITA = "(" & Mid(CADENITA, 1, Len(CADENITA) - 4) & ")"
          pub_cadena = pub_cadena + " AND  " & CADENITA
        End If
        
    Else
       If pub_cadena = "" Then
           pub_cadena = pub_cadena + " {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' "
       Else
            pub_cadena = pub_cadena + " AND  {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' "
       End If
    End If
    If LK_FLAG_SOS = "A" Then
       If pub_cadena = "" Then
           pub_cadena = pub_cadena + " {FACART.FAR_FLAG_SO} = 'A' "
       Else
           pub_cadena = pub_cadena + " AND {FACART.FAR_FLAG_SO} = 'A' "
       End If
    End If
  End If
  If tra_llave!tra_con14 = 1 Then ' x FECHAS X ALLOG
    If Not SON_FECHAS Then
     Exit Sub
    End If
    DIA = Day(REP_FECHA1)
    MES = Month(REP_FECHA1)
    ano = Year(REP_FECHA1)
    DIA1 = Day(REP_FECHA2)
    MES1 = Month(REP_FECHA2)
    ANO1 = Year(REP_FECHA2)
    CADENITA = "{ALLOG.ALL_FECHA_DIA} >= Date ( " & ano & "," & MES & "," & DIA & ") AND {ALLOG.ALL_FECHA_DIA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
    If pub_cadena = "" Then
       pub_cadena = pub_cadena + CADENITA
    Else
      pub_cadena = pub_cadena + " AND " + CADENITA
    End If
    If pub_cadena = "" Then
         pub_cadena = pub_cadena + " {ALLOG.ALL_CODCIA} = '" & LK_CODCIA & "' "
    Else
         pub_cadena = pub_cadena + " AND  {ALLOG.ALL_CODCIA} = '" & LK_CODCIA & "' "
    End If
    If LK_FLAG_SOS = "A" Then
       If pub_cadena = "" Then
           pub_cadena = pub_cadena + " {ALLOG.ALL_FLAG_SO} = 'A' "
       Else
           pub_cadena = pub_cadena + " AND {ALLOG.ALL_FLAG_SO} = 'A' "
       End If
    End If
  End If
  If tra_llave!tra_con10 = 1 Then ' x FECHA DE CARTERA VCTO
    If Not SON_FECHAS Then
      GoTo SALE
     Exit Sub
    End If
    DIA = Day(REP_FECHA1)
    MES = Month(REP_FECHA1)
    ano = Year(REP_FECHA1)
    DIA1 = Day(REP_FECHA2)
    MES1 = Month(REP_FECHA2)
    ANO1 = Year(REP_FECHA2)
    CADENITA = "{CARTERA.CAR_FECHA_VCTO} >= Date ( " & ano & "," & MES & "," & DIA & ") AND {CARTERA.CAR_FECHA_VCTO} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
    If pub_cadena = "" Then
       pub_cadena = pub_cadena + CADENITA
    Else
      pub_cadena = pub_cadena + " AND " + CADENITA
    End If
    If pub_cadena = "" Then
         pub_cadena = pub_cadena + " {CARTERA.CAR_CODCIA} = '" & LK_CODCIA & "' "
    Else
         pub_cadena = pub_cadena + " AND  {CARTERA.CAR_CODCIA} = '" & LK_CODCIA & "' "
    End If
    If LK_FLAG_SOS = "A" Then
       If pub_cadena = "" Then
           pub_cadena = pub_cadena + " {CARTERA.CAR_FLAG_SO} = 'A' "
       Else
           pub_cadena = pub_cadena + " AND {CARTERA.CAR_FLAG_SO} = 'A' "
       End If
    End If
  End If
  If tra_llave!tra_act8 = 1 Then ' x FECHA DE CARTERA
    If Not SON_FECHAS Then
      GoTo SALE
     Exit Sub
    End If
    DIA = Day(REP_FECHA1)
    MES = Month(REP_FECHA1)
    ano = Year(REP_FECHA1)
    DIA1 = Day(REP_FECHA2)
    MES1 = Month(REP_FECHA2)
    ANO1 = Year(REP_FECHA2)
    CADENITA = "{CARTERA.CAR_FECHA_INGR} >= Date ( " & ano & "," & MES & "," & DIA & ") AND {CARTERA.CAR_FECHA_INGR} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
    If pub_cadena = "" Then
       pub_cadena = pub_cadena + CADENITA
    Else
      pub_cadena = pub_cadena + " AND " + CADENITA
    End If
    If pub_cadena = "" Then
         pub_cadena = pub_cadena + " {CARTERA.CAR_CODCIA} = '" & LK_CODCIA & "' "
    Else
         pub_cadena = pub_cadena + " AND  {CARTERA.CAR_CODCIA} = '" & LK_CODCIA & "' "
    End If
    If LK_FLAG_SOS = "A" Then
       If pub_cadena = "" Then
           pub_cadena = pub_cadena + " {CARTERA.CAR_FLAG_SO} = 'A' "
       Else
           pub_cadena = pub_cadena + " AND {CARTERA.CAR_FLAG_SO} = 'A' "
       End If
    End If
  End If
  If tra_llave!tra_con12 = 1 Then ' x FECHA DE CARACU
    If Not SON_FECHAS Then
     Exit Sub
    End If
    DIA = Day(REP_FECHA1)
    MES = Month(REP_FECHA1)
    ano = Year(REP_FECHA1)
    DIA1 = Day(REP_FECHA2)
    MES1 = Month(REP_FECHA2)
    ANO1 = Year(REP_FECHA2)
    CADENITA = "{CARACU.CAA_FECHA} >= Date ( " & ano & "," & MES & "," & DIA & ") AND {CARACU.CAA_FECHA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
    If pub_cadena = "" Then
       pub_cadena = pub_cadena + CADENITA
    Else
      pub_cadena = pub_cadena + " AND " + CADENITA
    End If
    If tra_llave!tra_ACT11 = 0 Then
     If pub_cadena = "" Then
          pub_cadena = pub_cadena + " {CARACU.CAA_CODCIA} = '" & LK_CODCIA & "' "
     Else
          pub_cadena = pub_cadena + " AND  {CARACU.CAA_CODCIA} = '" & LK_CODCIA & "' "
     End If
    End If
    If LK_FLAG_SOS = "A" Then
       If pub_cadena = "" Then
           pub_cadena = pub_cadena + " {CARACU.CAA_FLAG_SO} = 'A' "
       Else
           pub_cadena = pub_cadena + " AND {CARACU.CAA_FLAG_SO} = 'A' "
       End If
    End If
    
  End If
  If tra_llave!TRA_CON11 = 1 Then
     CADENITA = " {CLIENTES.CLI_CODCIA} = '" & LK_CODCIA & "' AND {CLIENTES.CLI_CP} = '" & Left(cmbclipro.Text, 1) & "'"
     If pub_cadena = "" Then
         pub_cadena = pub_cadena + CADENITA
     Else
        If CADENITA <> "" Then
         pub_cadena = pub_cadena + " AND " + CADENITA
        End If
     End If
  End If
  If tra_llave!TRA_CON3 = 1 Then ' x VENDEDOR
      GoSub ARMA_VEND:
      If pub_cadena = "" Then
         pub_cadena = pub_cadena + CADENITA
      Else
        If CADENITA <> "" Then
         pub_cadena = pub_cadena + " AND " + CADENITA
        End If
      End If
      If pub_cadena = "" Then
         pub_cadena = pub_cadena + " {VEMAEST.VEM_CODCIA} = '" & LK_CODCIA & "' "
      Else
         pub_cadena = pub_cadena + " AND  {VEMAEST.VEM_CODCIA} = '" & LK_CODCIA & "' "
      End If
      wmensa = wmensa + "Ven.: " + wfiltra
  End If
  If tra_llave!TRA_ACT4 = 1 Then ' x VENDEDOR
      GoSub ARMA_TIPDOC
      If pub_cadena = "" Then
         pub_cadena = pub_cadena + CADENITA
      Else
        If CADENITA <> "" Then
         pub_cadena = pub_cadena + " AND " + CADENITA
        End If
      End If
      GoSub ARMA_SITUACION
      If pub_cadena = "" Then
         pub_cadena = pub_cadena + CADENITA
      Else
        If CADENITA <> "" Then
         pub_cadena = pub_cadena + " AND " + CADENITA
        End If
      End If
      
  End If
  If tra_llave!TRA_ACT6 = 1 Then
    If pub_cadena = "" Then
       pub_cadena = "{ARTI.ART_MONEDA} = '" & Left(cmdmoneda.Text, 1) & "'"
    Else
        pub_cadena = pub_cadena + " AND " + "{ARTI.ART_MONEDA} = '" & Left(cmdmoneda.Text, 1) & "'"
    End If
  End If
  If tra_llave!TRA_ACT7 = 1 Then
    If pub_cadena = "" Then
          If cheestado.Value = 0 Then
             pub_cadena = "{CLIENTES.CLI_ESTADO} = 'A'"
          End If
    Else
          If cheestado.Value = 0 Then
             pub_cadena = pub_cadena + " AND {CLIENTES.CLI_ESTADO} = 'A'"
          End If
    End If
  
  End If
  If tra_llave!TRA_ACT14 = 1 Then
    WFECHA_DIA = Format(LK_FECHA_DIA, "dd/mm/") & Format((Val(Format(LK_FECHA_DIA, "yyyy")) - 6), "####")
  Else
    WFECHA_DIA = Format(LK_FECHA_DIA, "dd/mm/yyyy")
  End If
  If tra_llave!tra_ACT1 = 1 Then
    wf1 = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
  End If
  If tra_llave!tra_ACT2 = 1 Then
    wf2 = "DIA=  '" & WFECHA_DIA & "'"
  End If
  If tra_llave!tra_ACT10 = 1 Then
    wf3 = "FECHAS=  ' DEL " & REP_FECHA1 & " AL " & REP_FECHA2 & "'"
  End If
  If tra_llave!tra_con15 = 1 Then ' X para fecha de rango en columnas 12 maximos
    GoSub PRO_COLU
  End If
  If wf1 <> "" Then Reportes.Formulas(0) = wf1
  If wf2 <> "" Then Reportes.Formulas(1) = wf2
  If wf3 <> "" Then Reportes.Formulas(2) = wf3
  If wf4 <> "" Then Reportes.Formulas(3) = wf4
  If wf5 <> "" Then Reportes.Formulas(4) = wf5
  If wf6 <> "" Then Reportes.Formulas(5) = wf6
  If wf7 <> "" Then Reportes.Formulas(6) = wf7
  If wf8 <> "" Then Reportes.Formulas(7) = wf8
  If wf9 <> "" Then Reportes.Formulas(8) = wf9
  If wf10 <> "" Then Reportes.Formulas(9) = wf10
  
  Reportes.Formulas(50) = ""
  If tra_llave!tra_ACT3 = 1 Then
      DIA = Day(LK_FECHA_DIA)
      MES = Month(LK_FECHA_DIA)
      ano = Year(LK_FECHA_DIA)
      Reportes.Formulas(50) = "FECHADIA= Date ( " & ano & "," & MES & "," & DIA & ")"
  End If
  Reportes.SelectionFormula = pub_cadena

  'Debug.Print pub_cadena
  Reportes.Action = 1
  ProgBar.Value = 10
  Screen.MousePointer = 0
  ProgBar.Visible = False
  lblproceso.Visible = False
  Pantalla.Enabled = True
  cmdcerrar.Enabled = True

Exit Sub

ARMA_ZONA:
Dim WTIPREG As Integer
Dim ALIAS_TABLAS As String
CADENITA = ""
wfiltra = ""
If opzonas(0).Value Then
 Modo1 = "{CLIENTES.CLI_CASA_ZONA} in ["
 ALIAS_TABLAS = "{ZONAS.TAB_TIPREG} = "
 WTIPREG = 20
ElseIf opzonas(1).Value Then
 Modo1 = "{CLIENTES.CLI_CASA_SUBZONA} in ["
 ALIAS_TABLAS = "{SUB_ZONAS.TAB_TIPREG} ="
 WTIPREG = 30
ElseIf opzonas(2).Value Then
 Modo1 = "{CLIENTES.CLI_ZONA_NEW} in ["
 ALIAS_TABLAS = "{ZONA_NEW.TAB_TIPREG} ="
 
 WTIPREG = 35
Else
GoTo pasa
End If
For fila = 0 To zonas.ListCount - 1
  zonas.ListIndex = fila
  If zonas.Selected(fila) Then
    wfiltra = wfiltra + Left(zonas.Text, 8) + ","
    Modo1 = Modo1 + Str(Val(Right(zonas.Text, 6))) + ","
  End If
Next fila
If wfiltra <> "" Then
  CADENITA = ALIAS_TABLAS & WTIPREG & " AND " & Left(Modo1, Len(Modo1) - 1) & "] "
  wfiltra = Left(wfiltra, Len(wfiltra) - 1)
Else
  CADENITA = ALIAS_TABLAS & WTIPREG & ""
  wfiltra = "(*)"
End If
pasa:
Return

ARMA_FAMI:
CADENITA = ""
wfiltra = ""
If Nulo_Valor0(tra_llave!TRA_CON13) = 1 Then
  Modo1 = "{FAMILIA.TAB_NUMTAB} in ["
Else
  Modo1 = "{ARTI.ART_FAMILIA} in ["
End If
For fila = 0 To fami.ListCount - 1
  fami.ListIndex = fila
  If fami.Selected(fila) Then
    wfiltra = wfiltra + Left(fami.Text, 8) + ","
    Modo1 = Modo1 + Str(Val(Right(fami.Text, 6))) + ","
  End If
Next fila
If wfiltra <> "" Then
  CADENITA = Left(Modo1, Len(Modo1) - 1) & "] "
  wfiltra = Left(wfiltra, Len(wfiltra) - 1)
Else
  CADENITA = ""
  wfiltra = "(*)"
End If

If Nulo_Valor0(tra_llave!TRA_CON13) = 1 Then
  If CADENITA <> "" Then
     CADENITA = CADENITA + " AND {FAMILIA.TAB_TIPREG} = 122 AND {FAMILIA.TAB_CODCIA} = '" & wcodcia & "' "
  Else
     CADENITA = "{FAMILIA.TAB_TIPREG} = 122 AND {FAMILIA.TAB_CODCIA} = '" & wcodcia & "' "
  End If
End If

Return


ARMA_NUMERO:
CADENITA = ""
wfiltra = ""
Modo1 = "{ARTI.ART_NUMERO} in ["
For fila = 0 To art_numero.ListCount - 1
  art_numero.ListIndex = fila
  If art_numero.Selected(fila) Then
    wfiltra = wfiltra + Left(art_numero.Text, 8) + ","
    Modo1 = Modo1 + Str(Val(Right(art_numero.Text, 6))) + ","
  End If
Next fila
If wfiltra <> "" Then
  CADENITA = Left(Modo1, Len(Modo1) - 1) & "] "
  wfiltra = Left(wfiltra, Len(wfiltra) - 1)
Else
  CADENITA = ""
  wfiltra = "(*)"
End If

Return

ARMA_SUBFAMI:
CADENITA = ""
wfiltra = ""
If Nulo_Valor0(tra_llave!TRA_CON13) = 1 Then
  Modo1 = "{FAMILIA.TAB_CODCIA}= '" & wcodcia & "' AND {FAMILIA.TAB_TIPREG}= 122 AND {FAMILIA.TAB_NUMTAB} = " & Str(Val(Right(famix.Text, 6))) & " AND  {SUBFAM.TAB_NUMTAB} in ["
Else
  Modo1 = "{ARTI.ART_SUBFAM} in ["
End If


For fila = 0 To subfami.ListCount - 1
  subfami.ListIndex = fila
  If subfami.Selected(fila) Then
    wfiltra = wfiltra + Left(subfami.Text, 8) + ","
    Modo1 = Modo1 + Str(Val(Right(subfami.Text, 6))) + ","
  End If
Next fila
If wfiltra <> "" Then
  CADENITA = Left(Modo1, Len(Modo1) - 1) & "] "
  wfiltra = Left(wfiltra, Len(wfiltra) - 1)
Else
  CADENITA = ""
  wfiltra = "(*)"
End If
If Nulo_Valor0(tra_llave!TRA_CON13) = 1 Then
  If CADENITA <> "" Then
    CADENITA = CADENITA + " AND {SUBFAM.TAB_TIPREG} = 123 AND {SUBFAM.TAB_CODCIA} = '" & wcodcia & "' "
  Else
    Modo1 = "{FAMILIA.TAB_CODCIA}= '" & wcodcia & "' AND {FAMILIA.TAB_TIPREG}= 122 AND {FAMILIA.TAB_NUMTAB} = " & Str(Val(Right(famix.Text, 6))) & " AND "
    CADENITA = Modo1 & " {SUBFAM.TAB_TIPREG} = 123 AND {SUBFAM.TAB_CODCIA} = '" & wcodcia & "' "
  End If
End If

Return

ARMA_VEND:
CADENITA = ""
wfiltra = ""
Modo1 = "{VEMAEST.VEM_CODVEN} in ["
For fila = 0 To multiven.ListCount - 1
  multiven.ListIndex = fila
  If multiven.Selected(fila) Then
    wfiltra = wfiltra + Left(multiven.Text, 3) + ","
    Modo1 = Modo1 + Str(Val(Left(multiven.Text, 3))) + ","
  End If
Next fila
If wfiltra <> "" Then
  CADENITA = Left(Modo1, Len(Modo1) - 1) & "] "
  wfiltra = Left(wfiltra, Len(wfiltra) - 1)
Else
  CADENITA = ""
  wfiltra = "(*)"
End If
Return

ARMA_TIPDOC:
CADENITA = ""
wfiltra = ""
Modo1 = "{CARTERA.CAR_TIPDOC} in ["
For fila = 0 To TIPDOC.ListCount - 1
  TIPDOC.ListIndex = fila
  If TIPDOC.Selected(fila) Then
    wfiltra = wfiltra + Left(TIPDOC.Text, 2) + ","
    Modo1 = Modo1 + "'" + Left(TIPDOC.Text, 2) + "' ,"
  End If
Next fila
If wfiltra <> "" Then
  CADENITA = Left(Modo1, Len(Modo1) - 1) & "] "
Else
  CADENITA = ""
End If
Return

ARMA_SITUACION:
CADENITA = ""
wfiltra = ""
Modo1 = "{CARTERA.CAR_SITUACION} in ["
For fila = 0 To SITUACION.ListCount - 1
  SITUACION.ListIndex = fila
  If SITUACION.Selected(fila) Then
    wfiltra = wfiltra + Left(SITUACION.Text, 1) + ","
    Modo1 = Modo1 + "'" + Left(SITUACION.Text, 1) + "' ,"
  End If
Next fila
If wfiltra <> "" Then
  CADENITA = Left(Modo1, Len(Modo1) - 1) & "] "
Else
  CADENITA = ""
End If
Return


PRO_COLU:
Dim i As Integer
Dim xcuenta As Integer
Dim cm As Integer
Dim fec2 As Date

For fila = 1 To 50
 Reportes.Formulas(fila) = ""
Next fila
cm = DateDiff("m", REP_FECHA1, REP_FECHA2)


MES = Month(REP_FECHA1)
MES1 = Month(REP_FECHA2)
ano = Year(REP_FECHA1)
ANO1 = Year(REP_FECHA2)
If ano = ANO1 Then
  Reportes.Formulas(11) = "ANO = '" & ano & "'"
Else
  Reportes.Formulas(11) = "ANO = '" & ano & " - " & ANO1 & "'"
End If
If cm > 12 Then
 MES1 = MES + 11
Else
 MES1 = MES + cm
End If
'If (MES1 - MES) > 0 Then
'  MES1 = MES1 - (MES1 - MES)
'End If
'fec1 = REP_FECHA1
'fec2 = REP_FECHA2
'Do Until fec1 >= fec2
' fec1 = DateAdd("m", i, fec1)
'fec1 = DatePart
'Loop

xcuenta = 0
i = 1
For fila = MES To MES1
 If fila > 12 Then
    Reportes.Formulas(12 + xcuenta) = "M" & i & "=" & fila - 12
    xcuenta = xcuenta + 1
    Reportes.Formulas(12 + xcuenta) = "A" & i & "=" & ANO1
 Else
    Reportes.Formulas(12 + xcuenta) = "M" & i & "=" & fila
    xcuenta = xcuenta + 1
    Reportes.Formulas(12 + xcuenta) = "A" & i & "=" & ano
 End If
 xcuenta = xcuenta + 1
 i = i + 1
Next fila

Return




SALE:
 Screen.MousePointer = 0
 ProgBar.Visible = False
 lblproceso.Visible = False
 If Err.Number = 20504 Then
   MsgBox "el Informe no se encontro Verificar :" & Reportes.ReportFileName, 48, Pub_Titulo
 ElseIf Err.Number = 20510 Then
   MsgBox "Falta Crear alguna Formula en Informe Verificar ", 48, Pub_Titulo
 ElseIf Err.Number = 20515 Then
   MsgBox "Selección de información No procede. Verificar ", 48, Pub_Titulo
 Else
   MsgBox Err.Description & " .Verificar", 48, Pub_Titulo
 End If
 Pantalla.Enabled = True
 cmdcerrar.Enabled = True

End Sub
Private Sub txt_cli_GotFocus()
Azul txt_cli, txt_cli
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
   pu_cp = loc_cp
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
   If Pantalla.Visible And Pantalla.Enabled Then
     Pantalla.SetFocus
   End If
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
   If Pantalla.Visible And Pantalla.Enabled Then
     Pantalla.SetFocus
   End If
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
    archi = "SELECT CLI_CODCLIE, CLI_CODCIA, CLI_CP, CLI_NOMBRE,CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM  FROM CLIENTES WHERE  CLI_CP = '" & loc_cp & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & txt_cli.Text & "' AND  '" & var & "' ORDER BY CLI_NOMBRE"
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

Private Sub txtCampo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 txtCampo2.SetFocus
End If
End Sub

Public Sub LLENA_VENDEDORES()
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset
Dim codi As String * 3
 pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = ? ORDER BY VEM_CODVEN"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 PS_REP01(0) = 0
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
 PS_REP01(0) = LK_CODCIA
 llave_rep01.Requery
 multiven.Clear
 Do Until llave_rep01.EOF
     codi = Format(llave_rep01!vem_codven, "000")
     multiven.AddItem codi & " " & Trim(llave_rep01!VEM_NOMBRE)
     llave_rep01.MoveNext
 Loop
 multiven.Visible = True
End Sub

Private Sub txtcampo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Pantalla.Enabled Then Pantalla.SetFocus
End If
End Sub

Private Sub txt_key_GotFocus()
 Azul txt_key, txt_key
End Sub
Private Sub txt_key_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not ListView3.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And txt_key.Text = "" Then
  loc_key = 1
  Set ListView3.SelectedItem = ListView3.ListItems(loc_key)
  ListView3.ListItems.Item(loc_key).Selected = True
  ListView3.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView3.ListItems.Count Then loc_key = ListView3.ListItems.Count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView3.ListItems.Count Then loc_key = ListView3.ListItems.Count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
  ListView3.ListItems.Item(loc_key).Selected = True
  ListView3.ListItems.Item(loc_key).EnsureVisible
  txt_key.Text = Trim(ListView3.ListItems.Item(loc_key).Text) & " "
  txt_key.SelStart = Len(txt_key.Text)
fin:

End Sub
Private Sub txt_key_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As ListItem
'On Error GoTo SALCODI
If KeyAscii = 27 Then
 txt_key.Text = ""
End If
If KeyAscii <> 13 Then Exit Sub
pu_codclie = Val(txt_key.Text)
If Len(txt_key.Text) = 0 Then
   Exit Sub
End If
'fra2.Refresh
If pu_codclie <> 0 And IsNumeric(txt_key.Text) = True Then
    SQ_OPER = 1
    On Error GoTo mucho
    PUB_CODBAN = Val(txt_key.Text)
    On Error GoTo 0
    pu_codcia = LK_CODCIA
    LEER_CCM_LLAVE
    If ccm_llave.EOF Then
            MsgBox "Registro ,   NO EXISTE ... "
            Azul txt_key, txt_key
            GoTo fin
    End If
    lblbanco.Caption = Trim(ccm_llave!CCM_nombre)
    txt_key.Text = Trim(ccm_llave!CCM_CODBAN)
    If Pantalla.Visible And Pantalla.Enabled Then
      Pantalla.SetFocus
    End If
    ListView3.Visible = False

    Screen.MousePointer = 0
Else
   If loc_key > ListView3.ListItems.Count Or loc_key = 0 Then
     Exit Sub
   End If
   valor = UCase(ListView3.ListItems.Item(loc_key).Text)
   If Trim(UCase(txt_key.Text)) = Left(valor, Len(Trim(txt_key.Text))) Then
   Else
      Exit Sub
   End If
   lblbanco.Caption = Trim(ListView3.ListItems.Item(loc_key).Text)
   txt_key.Text = Trim(ListView3.ListItems.Item(loc_key).SubItems(1))
   If Pantalla.Visible And Pantalla.Enabled Then
     Pantalla.SetFocus
   End If

   ListView3.Visible = False
   
End If
dale:
ListView3.Visible = False
fin:
mucho:

Exit Sub
SALCODI:
MsgBox Err.Description & " Intente Nuevamente ", 48, Pub_Titulo

End Sub

Private Sub txt_key_KeyUp(KeyCode As Integer, Shift As Integer)
Dim NADA
Dim var
If Len(txt_key.Text) = 0 Or IsNumeric(txt_key.Text) = True Then
   ListView3.Visible = False
   Exit Sub
End If
If ListView3.Visible = False And KeyCode <> 13 Or Len(txt_key.Text) = 1 Then
    If txt_key.Text = "" Then txt_key.Text = " "
    var = Asc(txt_key.Text)
    var = var + 1
    NADA = var
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    numarchi = 0
    archi = "SELECT * FROM CCMAEST WHERE  CCM_CODCIA = '" & LK_CODCIA & "' AND CCM_NOMBRE BETWEEN '" & txt_key.Text & "' AND  '" & var & "' ORDER BY CCM_NOMBRE"
    PROC_LISVIEW ListView3
    loc_key = 1
    If NADA = 33 Or NADA = 91 Then
      If ListView3.Visible = False Then
        loc_key = 0
        MsgBox "No existe Datos ...", 48, Pub_Titulo
        txt_key.Text = ""
      End If
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If ListView3.Visible Then
  Set itmFound = ListView3.FindItem(LTrim(txt_key.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView3.ListItems.Count Then
      ListView3.ListItems.Item(ListView3.ListItems.Count).EnsureVisible
   Else
     ListView3.ListItems.Item(loc_key + 8).EnsureVisible
   End If
  End If
  Exit Sub
End If
End Sub

Private Sub ListView3_DblClick()
 loc_key = ListView3.SelectedItem.Index
 txt_key.Text = Trim(ListView3.ListItems.Item(loc_key).Text) & " "
 txt_key_KeyPress 13
End Sub

Private Sub ListView3_GotFocus()
If loc_key <> 0 Then
 Set ListView3.SelectedItem = ListView3.ListItems(loc_key)
 ListView3.ListItems.Item(loc_key).Selected = True
 ListView3.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub ListView3_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = ListView3.SelectedItem.Index
 txt_key.Text = Trim(ListView3.ListItems.Item(loc_key).Text) & " "
End If

End Sub

Private Sub ListView3_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 ListView3.Visible = False
 txt_key.Text = ""
 txt_key.SetFocus
 Exit Sub
End If
If KeyAscii <> 13 Then
 Exit Sub
End If
ListView3_DblClick

End Sub

Private Sub ListView3_LostFocus()
ListView3.Visible = False
End Sub

