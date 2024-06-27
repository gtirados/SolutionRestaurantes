VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FORM_CONTA2 
   BackColor       =   &H00FAEFDA&
   Caption         =   "Ingreso de Vouchers"
   ClientHeight    =   6750
   ClientLeft      =   1500
   ClientTop       =   1140
   ClientWidth     =   8130
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6750
   ScaleWidth      =   8130
   Tag             =   "55"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdimp 
      Caption         =   "&Imprimir Voucher"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9810
      TabIndex        =   52
      Top             =   1530
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FAEFDA&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   0
      TabIndex        =   23
      Top             =   15
      Width           =   9495
      Begin VB.ComboBox voucher 
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
         ItemData        =   "FORM_CONTA2.frx":0000
         Left            =   120
         List            =   "FORM_CONTA2.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox txtdetalle 
         Height          =   300
         Left            =   3960
         TabIndex        =   42
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   1215
         Width           =   5295
      End
      Begin MSMask.MaskEdBox i_fecha2 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   1005
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.TextBox tc 
         Enabled         =   0   'False
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
         Left            =   2040
         TabIndex        =   4
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox i_glosa 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   780
         Width           =   5295
      End
      Begin VB.ComboBox i_plantilla 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   5295
      End
      Begin VB.ComboBox i_moneda 
         Enabled         =   0   'False
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
         ItemData        =   "FORM_CONTA2.frx":0004
         Left            =   2070
         List            =   "FORM_CONTA2.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox i_tipdoc 
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1560
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox i_numser 
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
         Left            =   2040
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   1560
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox i_numfac 
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
         Left            =   2040
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label LCODART 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plantilla :"
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
         Index           =   8
         Left            =   3960
         TabIndex        =   74
         Tag             =   "9999"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label LCODART 
         Caption         =   "[Insert]=T.C."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   7
         Left            =   1920
         TabIndex        =   50
         Tag             =   "9999"
         Top             =   1605
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label LCODART 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Movimiento.:"
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
         Index           =   6
         Left            =   120
         TabIndex        =   49
         Tag             =   "9999"
         Top             =   120
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle : "
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
         Left            =   3165
         TabIndex        =   43
         Top             =   1245
         Width           =   735
      End
      Begin VB.Label LCODART 
         Caption         =   "T.C.(S/.)"
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
         Index           =   1
         Left            =   1080
         TabIndex        =   30
         Tag             =   "9999"
         Top             =   1680
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label LCODART 
         Caption         =   "Nº Doc."
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
         Index           =   4
         Left            =   2295
         TabIndex        =   28
         Tag             =   "9999"
         Top             =   1650
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label LCODART 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Glosa :"
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
         Index           =   5
         Left            =   3165
         TabIndex        =   27
         Tag             =   "9999"
         Top             =   810
         Width           =   555
      End
      Begin VB.Label LCODART 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Emisión:"
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
         Left            =   120
         TabIndex        =   26
         Tag             =   "9999"
         Top             =   765
         Width           =   1485
      End
      Begin VB.Label LCODART 
         Caption         =   "Moneda:"
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   1920
         TabIndex        =   25
         Tag             =   "9999"
         Top             =   1530
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label LCODART 
         Caption         =   "TipoDoc."
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
         Index           =   2
         Left            =   2715
         TabIndex        =   24
         Tag             =   "9999"
         Top             =   1590
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Opciones"
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
      Height          =   1950
      Left            =   9525
      TabIndex        =   17
      Top             =   0
      Width           =   2250
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Consultas x &Doc."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Consultas x &Voucher"
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
         Index           =   1
         Left            =   135
         TabIndex        =   21
         Top             =   645
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FAEFDA&
         Caption         =   "&Ingresos"
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
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Frame fra_cc 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Centro de Costos: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2595
      Left            =   5595
      TabIndex        =   75
      Top             =   3150
      Visible         =   0   'False
      Width           =   3600
      Begin VB.ComboBox lstcc 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Index           =   2
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   1950
         Width           =   2955
      End
      Begin VB.ComboBox lstcc 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Index           =   1
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   1335
         Width           =   2955
      End
      Begin VB.ComboBox lstcc 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Index           =   0
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   735
         Width           =   2955
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   " División - Grupo"
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
         Height          =   180
         Index           =   2
         Left            =   210
         TabIndex        =   81
         Top             =   1710
         Width           =   2955
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   " Tipo Movimiento"
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
         Height          =   180
         Index           =   1
         Left            =   210
         TabIndex        =   80
         Top             =   1110
         Width           =   2955
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   " Zona - Sucursal"
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
         Height          =   180
         Index           =   0
         Left            =   210
         TabIndex        =   79
         Top             =   495
         Width           =   2955
      End
   End
   Begin VB.ListBox liscc 
      Height          =   2985
      Left            =   3960
      TabIndex        =   71
      Top             =   3120
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Frame frmcon 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Consultas"
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   7485
      TabIndex        =   35
      Top             =   6465
      Visible         =   0   'False
      Width           =   3975
      Begin VB.ComboBox TIPMOV 
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
         ItemData        =   "FORM_CONTA2.frx":0018
         Left            =   1560
         List            =   "FORM_CONTA2.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   300
         Width           =   975
      End
      Begin VB.TextBox tvou 
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
         Left            =   2640
         TabIndex        =   38
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton cant 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   330
         Picture         =   "FORM_CONTA2.frx":0062
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   240
         Width           =   420
      End
      Begin VB.CommandButton csig 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   840
         Picture         =   "FORM_CONTA2.frx":0764
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   240
         Width           =   420
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Vou :"
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
         Left            =   2640
         TabIndex        =   39
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.CommandButton SSALIR 
      Cancel          =   -1  'True
      Caption         =   "&Regresar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6480
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar PB 
      Height          =   240
      Left            =   60
      TabIndex        =   31
      Top             =   1695
      Visible         =   0   'False
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid_cons 
      Height          =   1215
      Left            =   615
      TabIndex        =   18
      Tag             =   "9999"
      Top             =   3420
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2143
      _Version        =   393216
      Rows            =   3
      FixedRows       =   2
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      GridColorFixed  =   15386515
      FocusRect       =   2
      HighLight       =   2
      GridLinesFixed  =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frmdocu 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Consultar"
      Height          =   1575
      Left            =   7200
      TabIndex        =   33
      Top             =   2625
      Visible         =   0   'False
      Width           =   3975
      Begin MSFlexGridLib.MSFlexGrid grid_docu 
         Height          =   1935
         Left            =   165
         TabIndex        =   34
         Tag             =   "9999"
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   3413
         _Version        =   393216
         Rows            =   3
         Cols            =   4
         BackColorBkg    =   16118252
         GridColorFixed  =   8421504
         FocusRect       =   2
         HighLight       =   2
         GridLinesFixed  =   1
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame ESTADO 
      BackColor       =   &H00FAEFDA&
      ForeColor       =   &H00000000&
      Height          =   5895
      Left            =   0
      TabIndex        =   9
      Tag             =   "100"
      Top             =   1950
      Width           =   11775
      Begin VB.ComboBox cboSunat 
         Height          =   315
         Left            =   4860
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   555
         Visible         =   0   'False
         Width           =   1530
      End
      Begin ComctlLib.ListView LV_CLI 
         Height          =   375
         Left            =   3480
         TabIndex        =   16
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
      Begin VB.TextBox TEXTOVAR 
         Appearance      =   0  'Flat
         BackColor       =   &H00FAEFDA&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   480
         TabIndex        =   73
         Text            =   "TEXTOVAR"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TEXTOVAR2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FAEFDA&
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   2160
         TabIndex        =   72
         Text            =   "TEXTOVAR2"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdorden 
         Caption         =   "Proceso Orden"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8730
         TabIndex        =   67
         Top             =   5325
         Visible         =   0   'False
         Width           =   1305
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
         TabIndex        =   29
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FAEFDA&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   4500
         Width           =   7305
         Begin VB.CommandButton cancelar 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Cancelar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4812
            Style           =   1  'Graphical
            TabIndex        =   83
            TabStop         =   0   'False
            Tag             =   "9999"
            Top             =   240
            Width           =   1050
         End
         Begin VB.CommandButton SALIR 
            Caption         =   "Ce&rrar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5985
            Style           =   1  'Graphical
            TabIndex        =   82
            TabStop         =   0   'False
            Tag             =   "9999"
            Top             =   240
            Width           =   1050
         End
         Begin VB.CommandButton destino 
            Appearance      =   0  'Flat
            Caption         =   "&Ver Destino"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3639
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   240
            Width           =   1050
         End
         Begin VB.CommandButton cmdeli 
            Appearance      =   0  'Flat
            Caption         =   "&Eliminar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2466
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   240
            Width           =   1050
         End
         Begin VB.CommandButton cmdConsultar 
            Appearance      =   0  'Flat
            Caption         =   "&Mostrar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1293
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   240
            Width           =   1050
         End
         Begin VB.CommandButton cmdIngreso 
            Caption         =   "&Grabar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   240
            Width           =   1050
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grid_fac 
         Height          =   3825
         Left            =   120
         TabIndex        =   8
         Tag             =   "9999"
         Top             =   405
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   6747
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   8388608
         ForeColorFixed  =   16777215
         GridColorFixed  =   15386515
         AllowBigSelection=   -1  'True
         FocusRect       =   2
         HighLight       =   2
         GridLinesFixed  =   1
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblinforme 
         BackColor       =   &H00F5F1EC&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   180
         TabIndex        =   70
         Top             =   4275
         Width           =   7080
      End
      Begin VB.Label lbldes 
         BackStyle       =   0  'Transparent
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
         Left            =   5160
         TabIndex        =   68
         Top             =   4200
         Width           =   2295
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
         Left            =   5040
         TabIndex        =   66
         Top             =   4560
         Width           =   2535
      End
      Begin VB.Label men 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6720
         TabIndex        =   51
         Top             =   120
         Width           =   2775
      End
      Begin VB.Label lbldif 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
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
         Left            =   10080
         TabIndex        =   45
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Dif.:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   9600
         TabIndex        =   44
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Voucher: "
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3480
         TabIndex        =   41
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lbl_nro_voucher 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4800
         TabIndex        =   40
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblperiodo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   120
         Width           =   3165
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
         Left            =   480
         TabIndex        =   13
         Top             =   4080
         Width           =   3015
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
         TabIndex        =   12
         Top             =   1560
         Width           =   1575
      End
   End
   Begin ComctlLib.ProgressBar Barra 
      Height          =   255
      Left            =   240
      TabIndex        =   10
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
   Begin Crystal.CrystalReport Reportes 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame fratc 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Ajuste Automatico de Tipo de Cambio"
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
      Height          =   990
      Left            =   2640
      TabIndex        =   53
      Top             =   6960
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   4320
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3120
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1680
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3120
         TabIndex        =   57
         Text            =   "Text1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label ltc 
         BackStyle       =   0  'Transparent
         Caption         =   "S/."
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   58
         Top             =   240
         Width           =   255
      End
      Begin VB.Label ltc 
         BackStyle       =   0  'Transparent
         Caption         =   "Dif."
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   6
         Left            =   4320
         TabIndex        =   65
         Top             =   240
         Width           =   495
      End
      Begin VB.Label ltc 
         BackStyle       =   0  'Transparent
         Caption         =   "S/."
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   63
         Top             =   600
         Width           =   255
      End
      Begin VB.Label ltc 
         BackStyle       =   0  'Transparent
         Caption         =   "$."
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   61
         Top             =   600
         Width           =   255
      End
      Begin VB.Label ltc 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   60
         Top             =   600
         Width           =   855
      End
      Begin VB.Label ltc 
         BackStyle       =   0  'Transparent
         Caption         =   "$."
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   56
         Top             =   240
         Width           =   255
      End
      Begin VB.Label ltc 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Provision :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FORM_CONTA2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WPASA As Boolean
Dim WSELE As String * 1
Dim F2 As Integer
Dim llave1
Public WW_CUENTA As String
Dim FLAG_SALIR As Integer
Dim loc_key As Long
Dim XX_CUENTA As String * 12
Dim fila As Integer
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

Dim cc_activo As rdoResultset
Dim PSCC_ACTIVO As rdoQuery
Dim cc_exist As rdoResultset
Dim PSCC_EXIST As rdoQuery
Dim Doc_exist As rdoResultset
Dim PSDoc_EXIST As rdoQuery
Dim CC1 As String
Dim CC2 As String
Dim CC3 As String
Dim CC As String
Option Explicit
Public Sub grabar_movicont()
Dim WTIPMOV
Dim ws_tot_debe As Currency
Dim ws_tot_haber As Currency
Dim w_dh As String * 1
Dim FLAG As Boolean
Dim wscadena As String
Dim ffecha1
Dim ffecha2

FLAG = False
ws_tot_debe = Val(Format(FORM_CONTA2.grid_fac.TextMatrix(1, 11), "0.00"))
ws_tot_haber = Val(Format(FORM_CONTA2.grid_fac.TextMatrix(1, 12), "0.00"))

If ws_tot_debe = 0 And ws_tot_haber = 0 Then
  pub_mensaje = "Anulacion de Documentos!!! ...   ¿Desea Continuar... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbYes Then
      GoTo SIG
  End If
  Exit Sub
End If
SIG:
For fila = 2 To FORM_CONTA2.grid_fac.Rows - 1
 If FORM_CONTA2.grid_fac.TextMatrix(fila, 1) <> "" Then
    SQ_OPER = 1
    PUB_CUENTA = FORM_CONTA2.grid_fac.TextMatrix(fila, 1)
    PUB_CODCIA = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
       MsgBox "Cuenta No Existe en su Plan Contable : " & PUB_CUENTA, 48, Pub_Titulo
       Exit Sub
    Else
       'bloqueado por mic para para q permita pasar los diferentes niveles
     If Val(com_llave!com_flag_afectacion) <> 1 Then
'mic If Val(com_llave!com_nivel) <> cop_llave!cop_nivel_max Then
        MsgBox "Cuenta No es Analitica, Verificar : " & PUB_CUENTA, 48, Pub_Titulo
        Exit Sub
       End If
    End If
    
    If (Trim(FORM_CONTA2.grid_fac.TextMatrix(fila, 4)) = "C" Or Trim(FORM_CONTA2.grid_fac.TextMatrix(fila, 4)) = "P") And (Val(TIPMOV.Text) = 1 Or Val(TIPMOV.Text) = 2) Then
        SQ_OPER = 10
        pu_cp = Trim(FORM_CONTA2.grid_fac.TextMatrix(fila, 4))
        pu_codcia = LK_CODCIA
        pu_codclie = Val(FORM_CONTA2.grid_fac.TextMatrix(fila, 3))
        LEER_CLI_LLAVE
        If cli_llave10.EOF Then
          MsgBox "Ingresar el codigo de Cliente/Proveedor o No Existe, Verificar ", 48, Pub_Titulo
          Exit Sub
        End If
    
    End If
End If
Next fila
If ws_tot_debe <> ws_tot_haber Then
   pub_mensaje = MsgBox("Voucher no cuadra Desea Registrar ... ?", 36)
   If pub_mensaje = vbNo Then
     grid_fac.SetFocus
     Exit Sub
   End If
End If
If Trim(FORM_CONTA2.i_moneda.Text) = "" Then
   MsgBox ("Falta Indicar Moneda... ?")
   Exit Sub
End If
If Trim(FORM_CONTA2.i_tipdoc.Text) = "" Then
   MsgBox ("Falta Indicar Tipo de Documento... ?")
   Exit Sub
End If
OTROVAR:

pub_cadena = "SELECT * FROM CONTROLL"
Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
If loc_voucher <> -1 Then
    WTIPMOV = voucher.ItemData(voucher.ListIndex)
    ffecha1 = Format(LK_FECHA_COP1, "dd/mm/yyyy")
    ffecha2 = Format(LK_FECHA_COP2, "dd/mm/yyyy")
    pub_cadena = "DELETE MOVICONT  WHERE MOV_CODCIA = '" & LK_CODCIA & "' AND (MOV_FECHA >=  '" & ffecha1 & "'  AND MOV_FECHA <=  '" & ffecha2 & "') AND MOV_NRO_MES = " & LK_NRO_MES & " AND MOV_NRO_VOUCHER = " & loc_voucher & " AND MOV_TIPMOV = " & WTIPMOV
    CN.Execute pub_cadena, rdExecDirect
End If

PSMOV_VOU.rdoParameters(0) = LK_CODCIA
PSMOV_VOU.rdoParameters(1) = LK_FECHA_COP1
PSMOV_VOU.rdoParameters(2) = LK_FECHA_COP2
PSMOV_VOU.rdoParameters(3) = LK_NRO_MES
PSMOV_VOU.rdoParameters(4) = voucher.ItemData(voucher.ListIndex)
VOU_MOV.Requery
If VOU_MOV.EOF Then
   ws_nro_voucher = 0
Else
   ws_nro_voucher = VOU_MOV!MOV_NRO_VOUCHER
End If
If loc_voucher <> -1 Then
 ws_nro_voucher = loc_voucher ' Val(lbl_nro_voucher.Caption)
Else
 ws_nro_voucher = ws_nro_voucher + 1
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
con_llave.Close
lbl_nro_voucher.Caption = Format(ws_nro_voucher, "########0.0")
fila = 1
SUM_D = 0
SUM_H = 0
For fila = 2 To grid_fac.Rows - 1
  grid_fac.TextMatrix(fila, 3) = ""
  grid_fac.TextMatrix(fila, 4) = ""
  grid_fac.TextMatrix(fila, 5) = ""
  grid_fac.TextMatrix(fila, 6) = ""
  grid_fac.TextMatrix(fila, 8) = ""
  grid_fac.TextMatrix(fila, 12) = ""
  grid_fac.TextMatrix(fila, 13) = ""
Next fila

grid_fac.Rows = 2
i_moneda.ListIndex = 0

If LOC_DIF_TC <> 0 Then
     pub_mensaje = "Asiento por Diferencia de T.C. " & Chr(13) & "Por :S/. " & LOC_DIF_TC & " " & Chr(13) & "¿Desea Continuar... ?"
     Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
     If Pub_Respuesta = vbNo Then
         GoTo dale
     End If
     voucher.ListIndex = 3
    If LOC_DIF_TC < 0 Then
        grid_fac.Rows = grid_fac.Rows + 1
        grid_fac.TextMatrix(grid_fac.Rows - 1, 1) = CUENTA_DIF
        grid_fac.TextMatrix(grid_fac.Rows - 1, 3) = Abs(LOC_DIF_TC)
        grid_fac.TextMatrix(grid_fac.Rows - 1, 5) = LOC_CODCLIE
        grid_fac.TextMatrix(grid_fac.Rows - 1, 7) = loc_cp
        grid_fac.TextMatrix(grid_fac.Rows - 1, 8) = LOC_DOCU
        grid_fac.Rows = grid_fac.Rows + 1
        grid_fac.TextMatrix(grid_fac.Rows - 1, 1) = Trim(cop_llave!COP_CTA_DIF_TC_CONTRA)
        grid_fac.TextMatrix(grid_fac.Rows - 1, 4) = Abs(LOC_DIF_TC)
        grid_fac.TextMatrix(grid_fac.Rows - 1, 5) = ""
        grid_fac.TextMatrix(grid_fac.Rows - 1, 7) = ""
        grid_fac.TextMatrix(grid_fac.Rows - 1, 8) = ""
    Else
        grid_fac.Rows = grid_fac.Rows + 1
        grid_fac.TextMatrix(grid_fac.Rows - 1, 1) = CUENTA_DIF
        grid_fac.TextMatrix(grid_fac.Rows - 1, 4) = Abs(LOC_DIF_TC)
        grid_fac.TextMatrix(grid_fac.Rows - 1, 5) = LOC_CODCLIE
        grid_fac.TextMatrix(grid_fac.Rows - 1, 7) = loc_cp
        grid_fac.TextMatrix(grid_fac.Rows - 1, 8) = LOC_DOCU
        grid_fac.Rows = grid_fac.Rows + 1
        grid_fac.TextMatrix(grid_fac.Rows - 1, 1) = Trim(cop_llave!COP_CTA_DIF_TC_favor)
        grid_fac.TextMatrix(grid_fac.Rows - 1, 3) = Abs(LOC_DIF_TC)
        grid_fac.TextMatrix(grid_fac.Rows - 1, 5) = ""
        grid_fac.TextMatrix(grid_fac.Rows - 1, 9) = ""
        grid_fac.TextMatrix(grid_fac.Rows - 1, 8) = ""
    End If
    LOC_DIF_TC = 0
    GoTo OTROVAR
End If
dale:
Screen.MousePointer = 0
fila = 0
txtdetalle.Text = ""
TEXTOVAR.Text = ""
TEXTOVAR2.Visible = True
TEXTOVAR2.Text = ""
TEXTOVAR2.Visible = False
loc_voucher = -1
i_plantilla.SetFocus
SendKeys "%{up}"
Barra.Visible = False
cmdIngreso.Visible = False
cmdConsultar.Enabled = True

ACT_MESES (0)
Screen.MousePointer = 0




Exit Sub


ACT1:
fila = 1
PUB_FECHA = LK_FECHA_COP2
PUB_CONCEPTO = Trim(i_glosa.Text)
FLAG = False
WS_NRO_MOV = 1
fila = 2
Do While FLAG = False
   If Trim(FORM_CONTA2.grid_fac.TextMatrix(fila, 1)) = "" Then GoTo pasa
  ' PUB_FECHA = i_fecha2.Text
   PUB_CUENTA = Trim(FORM_CONTA2.grid_fac.TextMatrix(fila, 1))
   If Val(FORM_CONTA2.grid_fac.TextMatrix(fila, 11)) <> 0 Then
        w_dh = "D"
        PUB_IMPORTE = Val(FORM_CONTA2.grid_fac.TextMatrix(fila, 11))
   ElseIf Val(FORM_CONTA2.grid_fac.TextMatrix(fila, 12)) <> 0 Then
        w_dh = "H"
        PUB_IMPORTE = Val(FORM_CONTA2.grid_fac.TextMatrix(fila, 12))
   Else
      w_dh = "D"
      PUB_IMPORTE = 0
   End If
SIGUE_MAS:
    ' grabo todo
   temp_llave.AddNew
   temp_llave!MOV_CODCIA = LK_CODCIA
   temp_llave!MOV_PERIODO = Format(LK_FECHA_COP1, "yyyy")
   temp_llave!MOV_nro_MES = LK_NRO_MES
   temp_llave!MOV_NRO_VOUCHER = ws_nro_voucher
   temp_llave!MOV_NRO_MOV = WS_NRO_MOV
   temp_llave!MOV_TIPMOV = voucher.ItemData(voucher.ListIndex)
   
   temp_llave!MOV_FECHA = PUB_FECHA
   temp_llave!MOV_GLOSA = PUB_CONCEPTO
   temp_llave!MOV_MONEDA = Trim(i_moneda.Text)
   temp_llave!MOV_CODCTA = PUB_CUENTA
   temp_llave!MOV_DH = w_dh
   temp_llave!MOV_IMPORTE = PUB_IMPORTE
   temp_llave!MOV_SUNAT = Val(grid_fac.TextMatrix(fila, 7))
   temp_llave!MOV_serie = Val(grid_fac.TextMatrix(fila, 8))
   temp_llave!MOV_numfac = Val(grid_fac.TextMatrix(fila, 9))
   temp_llave!MOV_codclie = Val(grid_fac.TextMatrix(fila, 3))
   temp_llave!MOV_CP = Trim(grid_fac.TextMatrix(fila, 4))
  
   temp_llave!MOV_FBG = Trim(grid_fac.TextMatrix(fila, 7))
   temp_llave!MOV_MARCA = "X"
   temp_llave!MOV_DETALLE = Trim(txtdetalle.Text)
   temp_llave!MOV_FBG_C = " "
   temp_llave!MOV_numfac_c = 0
   temp_llave!MOV_serie_c = 0
   temp_llave!MOV_fecha_EMI = i_fecha2.Text
   temp_llave!MOV_PLANTILLA = Val(Right(i_plantilla.Text, 6))
   temp_llave!MOV_FLAG_TC = Trim(grid_fac.TextMatrix(fila, 25))
   temp_llave!MOV_TIPO_CAMBIO = Val(grid_fac.TextMatrix(fila, 24))
   temp_llave!MOV_FLAG_DES = " "
   temp_llave!MOV_CODUSU = LK_CODUSU
   temp_llave!MOV_CC = Trim(grid_fac.TextMatrix(fila, 6))
   temp_llave!MOV_OPC = Val(FORM_CONTA2.grid_fac.TextMatrix(fila, 5))
   temp_llave!MOV_EXONERADO = Val(FORM_CONTA2.grid_fac.TextMatrix(fila, 10))
   temp_llave!MOV_RUC = FORM_CONTA2.grid_fac.TextMatrix(fila, 20)
   temp_llave.Update
pasa:
   fila = fila + 1
   WS_NRO_MOV = WS_NRO_MOV + 1
   If fila >= FORM_CONTA2.grid_fac.Rows Then
      FLAG = True
   End If
  
Loop
If voucher.ItemData(voucher.ListIndex) = 2 Then
  i_numfac.Text = Val(FORM_CONTA2.i_numfac.Text) + 1
End If


cop_llave.Requery
cop_llave.Edit
cop_llave!COP_FLAG_MAYORIZACION = " "
cop_llave.Update
Return

Screen.MousePointer = 1


End Sub

Private Sub cancelar_Click()
'grid1.Visible = False
  grid_fac.BackColorFixed = &H800000
  grid_fac.GridColorFixed = &HEAC793
loc_voucher = -1
i_glosa.Text = ""
men.Caption = ""
'ESTADO.Caption = "Estado : "
fila = 0
SUM_D = 0
SUM_H = 0
LIMPIA_DATOS
CABE_ING
FORM_CONTA2.lcuenta.Caption = ""
'i_moneda.ListIndex = -1
cmdIngreso.Visible = False
cmdIngreso.Enabled = True
Grid_cons.Clear
Grid_cons.Cols = 1
Grid_cons.Rows = 1
Grid_cons.Visible = False
Option1(0).Value = True
voucher.SetFocus
SendKeys "%{UP}", True
TEXTOVAR.Visible = False
TEXTOVAR2.Visible = False
lbldif.Caption = ""
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
Exit Sub
SALE:
MsgBox "Verficar Importe.", 48, Pub_Titulo
If TEXTOVAR.Visible Then Azul3 TEXTOVAR, TEXTOVAR

End Sub

Private Sub cant_Click()
cant.Enabled = False
cant.Enabled = False
 If Val(tvou.Text) = 0 Then
    Exit Sub
 End If
 tvou.Text = Val(tvou.Text) - 1
 tvou_KeyPress 13
 cant.Enabled = True
 cant.Enabled = True

End Sub

Private Sub cboSunat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid_fac.TextMatrix(grid_fac.Row, 7) = Val(Right(cboSunat.Text, 10))
        cboSunat.Visible = False
        grid_fac.TextMatrix(grid_fac.Row, 22) = Left(cboSunat.Text, 30)
        grid_fac.Col = 8
    ElseIf KeyAscii = 27 Then
        cboSunat.Visible = False
        grid_fac.SetFocus
    End If
End Sub

Private Sub cmdConsultar_Click()
Dim flag_grabar As String * 1
Dim w_dh As String
Dim IMPORTE_DEB As Currency
Dim IMPORTE_HAB As Currency
Dim IMP_SUMA_D As Currency
Dim IMP_SUMA_H As Currency
Dim Wflag  As String * 1
Dim NumVoucher As Integer

If Grid_cons.Rows > 3 Then
' Grid_cons.Visible = True
' Grid_cons.SetFocus
' Exit Sub
End If
pub_mensaje = "Desea Mostar todos los Voucher... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If


loc_voucher = -1
If Left(cmdConsultar.Caption, 2) = "&d" Then
If Val(grid_fac.TextMatrix(1, 3)) = 0 Or Val(grid_fac.TextMatrix(1, 4)) = 0 Then Exit Sub
If grid_fac.Visible = False Then Exit Sub

  flag_grabar = ""
  Wflag = ""
  mov_llave.MoveFirst
  ws_nro_voucher = mov_llave!MOV_NRO_VOUCHER
  Do Until mov_llave.EOF
     mov_llave.Edit
     mov_llave.Delete
     mov_llave.MoveNext
  Loop
  grabar_movicont
  
  

  cop_llave.Requery
  cop_llave.Edit
  cop_llave!COP_FLAG_MAYORIZACION = " "
  cop_llave.Update
  MsgBox "Diario Actualizado. debe Mayorizar Nuevamente.", 48, Pub_Titulo
  cancelar_Click
  Exit Sub
End If

MA:
If Option1(2).Value = True Then
'   WW_CUENTA = InputBox("Ingrese Nro. Voucher Para la Consulta :")
'ElseIf Option1(0).Value = True Then
   WW_CUENTA = InputBox("Ingrese Numero de Documento")
   If WW_CUENTA = "" Then Exit Sub
End If
'If Option1(0).Value = 1 Or Option1(1).Value Then
'   If WW_CUENTA = "" Then GoTo MA
'End If



'Aqui empieza la consulta ......
ESTADO.Visible = False
SSALIR.Visible = True
Frame3.Visible = False
Frame2.Visible = False
Grid_cons.Visible = True

PSTEMP_LLAVE(0) = LK_CODCIA
PSTEMP_LLAVE(1) = LK_FECHA_COP1
PSTEMP_LLAVE(2) = LK_FECHA_COP2
temp_llave.Requery

ESTADO.Caption = "Estado :   < CONSULTA, MODIFICA, ELIMINA >"
CABE_MOSTRAR
i_glosa.Text = ""



WPASA = False


SIGUE:
Dim WS_SALDO As Currency
Dim Tit As String
Dim a As Integer
Dim i As Integer
Dim success%
Dim con_cuenta As String * 1

Screen.MousePointer = 11
Grid_cons.Visible = False
DoEvents
'success% = SetWindowPos(FrmProcesa.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
DoEvents
Dim ws_tot_debe, ws_tot_haber As Currency
Dim ws_fecha  As String
Dim ws_voucher As Currency



'PSTEMP_LLAVE.rdoParameters(0) = LK_CODCIA
'PSTEMP_LLAVE.rdoParameters(1) = i_fecha2.Text
'PSTEMP_LLAVE.rdoParameters(2) = i_fecha2.Text

'temp_llave.Requery
fila = 1
If temp_llave.EOF = True Then
   MsgBox "No hay Registros....", 48, Pub_Titulo
   'i_fecha2.SetFocus
   Screen.MousePointer = 0
   Grid_cons.Visible = True
   GoTo fin
End If
LOC_CANCELA = 1
'Frame1.Enabled = False
pb.Visible = True
DoEvents
pb.Min = 0
pb.Max = temp_llave.RowCount
pb.Value = 0
ws_tot_debe = 0
ws_tot_haber = 0
Tit = Grid_cons.FormatString
Grid_cons.FormatString = Tit
ws_fecha = temp_llave!MOV_fecha_EMI
ws_voucher = temp_llave!MOV_NRO_VOUCHER
Wflag = "X"
IMP_SUMA_D = 0
IMP_SUMA_H = 0
  '    Grid_cons.Visible = True
Do Until temp_llave.EOF
       pb.Value = pb.Value + 1
       If Option1(2) = True Then
          If Trim(temp_llave!MOV_numfac) <> Trim(WW_CUENTA) Then
 '             ws_voucher = temp_llave!MOV_nro_voucher
             GoTo OTRO
          End If
         End If
        If ws_voucher <> temp_llave!MOV_NRO_VOUCHER Then
          fila = fila + 1
          Grid_cons.Rows = fila + 1
          Grid_cons.TextMatrix(fila, 1) = "Totales "
          Grid_cons.TextMatrix(fila, 2) = ""
          Grid_cons.TextMatrix(fila, 3) = ""
          Grid_cons.TextMatrix(fila, 4) = Format(ws_tot_debe, "##,##0.00")
          Grid_cons.TextMatrix(fila, 5) = Format(ws_tot_haber, "##,##0.00")
          If ws_tot_debe <> ws_tot_haber Then MsgBox "Asiento no Cuadra Verificar  Voucher Nro: " & ws_voucher & " Fecha: " & Grid_cons.TextMatrix(fila - 1, 1) & " TipMov = " & NumVoucher, 48, Pub_Titulo
          Grid_cons.Row = fila
          Grid_cons.Col = 4
          Grid_cons.CellBackColor = vbCyan
          Grid_cons.Col = 5
          Grid_cons.CellBackColor = vbCyan
          ws_tot_debe = 0
          ws_tot_haber = 0
        End If
         
         'If Option1(1) = True Then
         ' If Trim(temp_llave!MOV_NRO_VOUCHER) <> Trim(WW_CUENTA) Then
         '    GoTo OTRO
         ' End If
         'End If
         
'          Grid_cons.Visible = True
       fila = fila + 1
       Grid_cons.Rows = fila + 1
       Grid_cons.TextMatrix(fila, 1) = Format(temp_llave!MOV_fecha_EMI, "dd-mm-yy")
       Grid_cons.TextMatrix(fila, 2) = Trim(temp_llave!MOV_CODCTA)
       
       'SQ_OPER = 1
       'PUB_CUENTA = Trim(temp_llave!MOV_CODCTA)
       'PUB_CODCIA = LK_CODCIA
       'LEER_COM_LLAVE
       'If Not com_llave.EOF Then
       ' Grid_cons.TextMatrix(fila, 3) = com_llave!com_DESCRIPCION
       'Else
       'MsgBox "Cuenta Verificar Cuenta : " & PUB_CUENTA
       'End If
       Grid_cons.TextMatrix(fila, 3) = "..."

       Grid_cons.TextMatrix(fila, 0) = temp_llave!MOV_NRO_VOUCHER
       NumVoucher = temp_llave!MOV_TIPMOV
       If temp_llave!MOV_DH = "D" Then
          Grid_cons.TextMatrix(fila, 4) = Nulo_Valor0(temp_llave!MOV_IMPORTE)
          Grid_cons.TextMatrix(fila, 5) = ""
          ws_tot_debe = ws_tot_debe + Nulo_Valor0(temp_llave!MOV_IMPORTE)
          IMP_SUMA_D = IMP_SUMA_D + temp_llave!MOV_IMPORTE
       ElseIf temp_llave!MOV_DH = "H" Then
             Grid_cons.TextMatrix(fila, 4) = ""
             Grid_cons.TextMatrix(fila, 5) = Nulo_Valor0(temp_llave!MOV_IMPORTE)
             ws_tot_haber = ws_tot_haber + Nulo_Valor0(temp_llave!MOV_IMPORTE)
             IMP_SUMA_H = IMP_SUMA_H + temp_llave!MOV_IMPORTE
       End If
       Grid_cons.TextMatrix(fila, 6) = Nulo_Valors(temp_llave!MOV_MONEDA)
       Grid_cons.TextMatrix(fila, 7) = Trim(temp_llave!MOV_SUNAT)
       Grid_cons.TextMatrix(fila, 8) = temp_llave!MOV_serie
       Grid_cons.TextMatrix(fila, 9) = temp_llave!MOV_numfac
       Grid_cons.TextMatrix(fila, 12) = Nulo_Valors(temp_llave!MOV_DETALLE)
       
OTRO:
       ws_fecha = temp_llave!MOV_fecha_EMI
       ws_voucher = temp_llave!MOV_NRO_VOUCHER
       temp_llave.MoveNext
       DoEvents
Loop
          fila = fila + 1
          Grid_cons.Rows = fila + 1
          Grid_cons.TextMatrix(fila, 1) = "Totales "
          Grid_cons.TextMatrix(fila, 2) = ""
          Grid_cons.TextMatrix(fila, 4) = Format(ws_tot_debe, "##,##0.00")
          Grid_cons.TextMatrix(fila, 5) = Format(ws_tot_haber, "##,##0.00")
          Grid_cons.Row = fila
          Grid_cons.Col = 4
          Grid_cons.CellBackColor = vbCyan
          Grid_cons.Col = 5
          Grid_cons.CellBackColor = vbCyan
If IMP_SUMA_H <> IMP_SUMA_D Then
  MsgBox "NO Cuadra el Total General del voucher", 48, Pub_Titulo
End If
Grid_cons.TextMatrix(1, 3) = "Tot. General ="
Grid_cons.TextMatrix(1, 4) = Format(IMP_SUMA_D, "0.00")
Grid_cons.TextMatrix(1, 5) = Format(IMP_SUMA_H, "0.00")
        
 LOC_CANCELA = 0
 cancelar.Enabled = True
 pb.Visible = False
 DoEvents
' suma_grid
 ws_tot_debe = 0
 ws_tot_haber = 0
 Grid_cons.Visible = True
 Frame1.Enabled = True
 cmdConsultar.Enabled = True
' Grid_cons.Col = 1
' Grid_cons.Row = 2
 If Grid_cons.Enabled And Grid_cons.Visible Then Grid_cons.SetFocus
 Screen.MousePointer = 0

fin:




End Sub

Private Sub cmdcorta_Click()
LOC_CANCELA = 2
End Sub

Private Sub cmdeli_Click()
Dim ffecha1
Dim ffecha2
Dim WTIPMOV
If Option1(1).Value Then
   If loc_voucher = -1 Then GoTo sa
    pub_mensaje = "Eliminar el Voucher!!! ...   ¿Desea Continuar... ? Nro. Voucher : " & loc_voucher
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
          Exit Sub
    End If
    ffecha1 = Format(LK_FECHA_COP1, "dd/mm/yyyy")
    ffecha2 = Format(LK_FECHA_COP2, "dd/mm/yyyy")
    WTIPMOV = TIPMOV.ListIndex + 1
    pub_cadena = "DELETE MOVICONT  WHERE MOV_CODCIA = '" & LK_CODCIA & "' AND MOV_FECHA >=  '" & ffecha1 & "'  AND MOV_FECHA <=  '" & ffecha2 & "' AND MOV_NRO_MES = " & LK_NRO_MES & " AND MOV_NRO_VOUCHER = " & loc_voucher & " AND MOV_TIPMOV = " & WTIPMOV
    CN.Execute pub_cadena, rdExecDirect
    cancelar_Click
Else
sa:
  MsgBox "Consulte antes de eliminar un voucher.", 48, Pub_Titulo
End If
End Sub

Private Sub cmdimp_Click()
On Error GoTo SALE
Dim CADENITA
Dim wvoucher As Integer
Dim WTIPMOV As Integer
WTIPMOV = voucher.ItemData(voucher.ListIndex)
wvoucher = Val(lbl_nro_voucher.Caption)
Reportes.ReportFileName = CONS_ADMIN & "CONTABILIDAD\" & "IMPVOU.RPT"
Reportes.Connect = PUB_ODBC
Reportes.WindowTitle = "Reporte :  VOUCHER "
Reportes.Destination = crptToWindow
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390
DoEvents
pub_cadena = "{MOVICONT.MOV_CODCIA} = '" & LK_CODCIA & "' and {MOVICONT.MOV_TIPMOV} = " & WTIPMOV & "  and {MOVICONT.MOV_NRO_MES} = " & LK_NRO_MES & " AND {MOVICONT.MOV_NRO_VOUCHER}  = " & wvoucher
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
Const modif = 1
Dim N As Integer
Dim LOC_SALDO_CAR As Currency
Dim FLAG As Boolean
Dim pub_mensaje_err As String
Dim w_dh  As String
Dim i As Integer
Dim wcadena As String
Dim wvalor  As String * 1
' CHEQUE DE MES CERRADO 0 ABIERTO

For i = 2 To grid_fac.Rows - 1
    If grid_fac.TextMatrix(i, 0) <> "" And (Trim(grid_fac.TextMatrix(i, 1)) = "") Then
        MsgBox "Error en los Datos"
        Exit Sub
    End If
Next i
wcadena = ""
SQ_OPER = 2
PUB_TIPREG = 155
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_mayor.EOF Then
  MsgBox "Crear tab_tipreg = 155 para seguridas", 48, Pub_Titulo
Else
    Do Until tab_mayor.EOF
      If tab_mayor!TAB_NUMTAB = Val(Format(LK_FECHA_COP1, "yyyy")) Then
        wcadena = Trim(tab_mayor!tab_nomlargo)
      End If
      tab_mayor.MoveNext
    Loop
    wvalor = Mid(wcadena, LK_NRO_MES + 1, 1)
    If wvalor = "1" Then
       MsgBox "<<<< Mes CERRADO Operaciones >>>>", vbCritical, Pub_Titulo
       Exit Sub
    End If
End If



OTRO:
cmdIngreso.Enabled = False
'MOV_NUMSER = ? AND MOV_NUMFAC = ? ORDER BY MOV_CODCIA, MOV_CODCTA , MOV_DH"
If voucher.ItemData(voucher.ListIndex) = 2 And Option1(0).Value Then
  PSPR_NUMFAC(0) = LK_CODCIA
  PSPR_NUMFAC(1) = voucher.ItemData(voucher.ListIndex)
  PSPR_NUMFAC(2) = LK_FECHA_COP1
  PSPR_NUMFAC(3) = LK_FECHA_COP2
  PSPR_NUMFAC(4) = Left(i_tipdoc.Text, 2)
  i_numser.Text = grid_fac.TextMatrix(2, 8)
  i_numfac.Text = grid_fac.TextMatrix(2, 9)
  PSPR_NUMFAC(5) = Val(i_numser.Text)
  PSPR_NUMFAC(6) = Val(i_numfac.Text)
  pr_numfac.Requery
  If Val(Left(i_tipdoc.Text, 2)) <> 0 Then
    If Not pr_numfac.EOF Then
      cmdIngreso.Enabled = True
      MsgBox "Existe Documento, Nro. voucher : ", 48, Pub_Titulo
      Exit Sub
    End If
  End If
End If
grabar_movicont
cmdIngreso.Enabled = True
TEXTOVAR.Visible = False
TEXTOVAR2.Visible = False
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
  FORM_CONTA2.Barra.Visible = False
  Screen.MousePointer = 0
  grid_fac.SetFocus
Else
  MsgBox Err.Description, 48, Pub_Titulo
End If

End Sub



Private Sub CmdOrden_Click()
Dim wfila As Integer
Dim wvoucher As Integer
pub_cadena = "SELECT MOV_NRO_MOV, MOV_NRO_VOUCHER FROM MOVICONT WHERE MOV_CODCIA = ? AND (MOV_FECHA >=? AND MOV_FECHA <=?)  AND MOV_TIPMOV = ? AND MOV_NRO_MES = " & LK_NRO_MES & "   ORDER BY MOV_TIPMOV, MOV_NRO_VOUCHER"
Set PSTEMP_LLAVE = CN.CreateQuery("", pub_cadena)
PSTEMP_LLAVE(0) = 0
PSTEMP_LLAVE(1) = LK_FECHA_DIA
PSTEMP_LLAVE(2) = LK_FECHA_DIA
PSTEMP_LLAVE(3) = 0
Set temp_llave = PSTEMP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
PSTEMP_LLAVE(0) = LK_CODCIA
PSTEMP_LLAVE(1) = LK_FECHA_COP1
PSTEMP_LLAVE(2) = LK_FECHA_COP2
PSTEMP_LLAVE(3) = voucher.ItemData(voucher.ListIndex)
temp_llave.Requery
If temp_llave.EOF Then
 MsgBox "no hay datos"
 Exit Sub
End If
wfila = 0
wvoucher = temp_llave!MOV_NRO_VOUCHER
Do Until temp_llave.EOF
 If wvoucher <> temp_llave!MOV_NRO_VOUCHER Then
  wfila = wfila + 1
  wvoucher = temp_llave!MOV_NRO_VOUCHER
 End If
 DoEvents
 cmdorden.Caption = temp_llave.RowCount & " " & wfila
 temp_llave.Edit
 temp_llave!MOV_NRO_MOV = wfila
 temp_llave.Update
 temp_llave.MoveNext
Loop
MsgBox " listo " & voucher.Text
cmdorden.Caption = "Ordernar Voucher"


End Sub

Private Sub csig_Click()
cant.Enabled = False
cant.Enabled = False
tvou.Text = Val(tvou.Text) + 1
tvou_KeyPress 13
cant.Enabled = True
cant.Enabled = True
End Sub

Private Sub destino_Click()
Dim went As Integer
Dim wvalor As Currency
went = Int(Val(tvou))
wvalor = Val(went) - Val(tvou.Text)
If wvalor <> 0 Then
  tvou.Text = Format(Val(tvou), "0")
  destino.Caption = "&Ver Destino"
  grid_fac.BackColorFixed = &H800000
  grid_fac.GridColorFixed = &HEAC793
Else
  tvou.Text = Format(Val(tvou) + 0.1, "0.0")
  destino.Caption = "Re&gresar"
  grid_fac.BackColorFixed = &HA5701F
  grid_fac.GridColorFixed = &H800000
End If

tvou_KeyPress 13

End Sub

Private Sub Form_Activate()
voucher.SetFocus
End Sub

Private Sub Form_DblClick()
pub_cadena = "SELECT * FROM MOVICONT WHERE MOV_CODCIA = ?  AND MOV_NRO_MES = " & LK_NRO_MES & "   ORDER BY MOV_TIPMOV ,MOV_NRO_VOUCHER, MOV_DH, MOV_NRO_MOV"
Set PSTEMP_LLAVE = CN.CreateQuery("", pub_cadena)
PSTEMP_LLAVE(0) = 0
'PSTEMP_LLAVE(1) = LK_FECHA_DIA
'PSTEMP_LLAVE(2) = LK_FECHA_DIA
Set temp_llave = PSTEMP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
PSTEMP_LLAVE(0) = LK_CODCIA
'PSTEMP_LLAVE(1) = LK_FECHA_COP1
'PSTEMP_LLAVE(2) = LK_FECHA_COP1
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If fra_cc.Visible = True And KeyCode = 27 Then
    fra_cc.Visible = False
End If
End Sub

Private Sub Form_Load()
'On Error GoTo SALE
    LlenaCboTablas cboSunat, 50, "00"
    
Wsec = 0
LOC_CANCELA = 0
fila = 0
wfila_act = 0
WSELE = ""
Dim ws_indice As Integer
Dim cade

pub_cadena = "SELECT MOV_FBG, MOV_NUMFAC, MOV_NRO_MES, MOV_NRO_VOUCHER, MOV_MONEDA, MOV_FECHA_EMI, MOV_SERIE, MOV_IMPORTE, MOV_DH , MOV_TIPMOV FROM MOVICONT WHERE MOV_CODCIA = ? AND MOV_CODCLIE = ? AND MOV_CP = ? AND MOV_FBG =  ? AND MOV_SERIE = ? AND MOV_NUMFAC = ? ORDER BY  MOV_FBG, MOV_NUMFAC"
Set PSDOCU = CN.CreateQuery("", pub_cadena)
PSDOCU(0) = 0
PSDOCU(1) = 0
PSDOCU(2) = 0
PSDOCU(3) = 0
PSDOCU(4) = 0
PSDOCU(5) = 0
Set leer_docu = PSDOCU.OpenResultset(rdOpenKeyset, rdConcurValues)


pub_cadena = "SELECT * FROM PLANTILLA WHERE PLT_CODCIA = ? AND PLT_TIPMOV = ? AND PLT_NUMERO >= ? ORDER BY  PLT_NUMERO, PLT_SECUENCIA"
Set PSTEMP_MAYOR = CN.CreateQuery("", pub_cadena)
PSTEMP_MAYOR(0) = 0
PSTEMP_MAYOR(1) = 0
PSTEMP_MAYOR(2) = 0
Set temp_mayor = PSTEMP_MAYOR.OpenResultset(rdOpenForwardOnly, rdConcurValues)


pub_cadena = "SELECT * FROM MOVICONT WHERE MOV_CODCIA = ? AND (MOV_FECHA >=? AND MOV_FECHA <=?)  AND MOV_NRO_MES = " & LK_NRO_MES & "   ORDER BY MOV_FECHA_EMI, MOV_NRO_VOUCHER, MOV_DH, MOV_NRO_MOV"
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

pub_cadena = "SELECT * FROM MOVICONT WHERE MOV_CODCIA = ?  AND MOV_NRO_VOUCHER =?  AND MOV_TIPMOV = ?  AND (MOV_FECHA >= ? AND MOV_FECHA <= ?) AND MOV_NRO_MES = " & LK_NRO_MES ' & " ORDER BY MOV_NRO_MOV"
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
 Unload FORM_CONTA2
 Exit Sub
End If
If LK_NRO_MES = 0 Then
   lblperiodo.Caption = "PERIODO : APERTURA"
Else
   lblperiodo.Caption = "PERIODO : " & Format(LK_FECHA_COP1, "dd/mm/yyyy") & " al " & Format(LK_FECHA_COP2, "dd/mm/yyyy")
End If

PUB_TIPREG = 50
PUB_CODCIA = "00"
SQ_OPER = 2
LEER_TAB_LLAVE
i_tipdoc.Clear
Do Until tab_mayor.EOF
   i_tipdoc.AddItem Format(tab_mayor!TAB_NUMTAB, "00") & ".-" & tab_mayor!tab_nomlargo
   tab_mayor.MoveNext
Loop
If i_tipdoc.ListCount > 0 Then i_tipdoc.ListIndex = 0

TIPMOV.Clear
PUB_TIPREG = 150
PUB_CODCIA = "00"
SQ_OPER = 2
LEER_TAB_LLAVE
Do Until tab_mayor.EOF
    voucher.AddItem Trim(tab_mayor!tab_nomlargo) & String(80, " ") & Trim(tab_mayor!TAB_CONTABLE2)
    voucher.ItemData(voucher.NewIndex) = tab_mayor!TAB_NUMTAB
    
    TIPMOV.AddItem tab_mayor!TAB_NUMTAB & " " & Trim(tab_mayor!tab_nomcorto) & String(80, " ") & Trim(tab_mayor!TAB_CONTABLE2)
    TIPMOV.ItemData(TIPMOV.NewIndex) = tab_mayor!TAB_NUMTAB
    
    tab_mayor.MoveNext
Loop

'=========================
'=========================
pub_cadena = "SELECT * FROM CENTROC WHERE CC_CODCIA = ? AND CC_TIPO = ? "
Set PSCC_ACTIVO = CN.CreateQuery("", pub_cadena)
PSCC_ACTIVO(0) = LK_CODCIA
PSCC_ACTIVO(1) = 1
If PUB_Flag_CC1 = 1 Then
    Set cc_activo = PSCC_ACTIVO.OpenResultset(rdOpenKeyset, rdConcurValues)
    Do While Not cc_activo.EOF
        lstcc(0).AddItem Trim(cc_activo("CC_DESCRIPCION")) 'Trim(cc_activo("CC_CODIGO")) + "-" +
        lstcc(0).ItemData(lstcc(0).NewIndex) = Val(cc_activo("CC_CODIGO"))
        cc_activo.MoveNext
    Loop
Else
    lstcc(0).Enabled = False
End If
If PUB_Flag_CC2 = 1 Then
    PSCC_ACTIVO(1) = 2
    Set cc_activo = PSCC_ACTIVO.OpenResultset(rdOpenKeyset, rdConcurValues)
    Do While Not cc_activo.EOF
        lstcc(1).AddItem Trim(cc_activo("CC_DESCRIPCION"))  'Trim(cc_activo("CC_CODIGO")) + "-" +
        lstcc(1).ItemData(lstcc(1).NewIndex) = Val(cc_activo("CC_CODIGO"))
        cc_activo.MoveNext
    Loop
Else
    lstcc(1).Enabled = False
End If
If PUB_Flag_CC3 = 1 Then
    PSCC_ACTIVO(1) = 3
    Set cc_activo = PSCC_ACTIVO.OpenResultset(rdOpenKeyset, rdConcurValues)
    Do While Not cc_activo.EOF
        lstcc(2).AddItem Trim(cc_activo("CC_DESCRIPCION")) 'Trim(cc_activo("CC_CODIGO")) + "-" +
        lstcc(2).ItemData(lstcc(2).NewIndex) = Val(cc_activo("CC_CODIGO"))
        cc_activo.MoveNext
    Loop
Else
    lstcc(2).Enabled = False
End If
pub_cadena = "SELECT * FROM Centroc WHERE (CC_TIPO = 1 AND CC_CODIGO = ?) OR (CC_TIPO = 2 AND CC_CODIGO = ?) OR (CC_TIPO = 3 AND CC_CODIGO = ?)"
Set PSCC_EXIST = CN.CreateQuery("", pub_cadena)
PSCC_EXIST(0) = 0
PSCC_EXIST(1) = 0
PSCC_EXIST(2) = 0
Set cc_exist = PSCC_EXIST.OpenResultset(rdOpenKeyset, rdConcurValues)

i_moneda.ListIndex = 0
'=========================
'=========================

'If LK_CODUSU = "ADMIN" Then
'  cmdorden.Visible = True
'End If
FLAG_DIF_TC = ""
fila = 0
DoEvents
LIMPIA_DATOS
Dim ws_fecha As Date
i_fecha2.Text = Format(LK_FECHA_COP2, "dd/mm/yyyy")
i_fecha2.Mask = "##/##/####"
fila = 0
Grid_cons.Top = 400
Grid_cons.Left = 200
Grid_cons.Width = grid_fac.Width - 500
Grid_cons.Height = 5000
Grid_cons.Cols = 13
Grid_cons.Visible = False
SSALIR.Visible = False
SSALIR.Top = 6400
SSALIR.Left = 8500
cmdIngreso.Visible = False
voucher.TabIndex = 0
loc_voucher = -1
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

Private Sub Grid_cons_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 Grid_cons.Visible = False
 Frame1.Enabled = True
 cmdIngreso.SetFocus
 Exit Sub
End If

If KeyAscii <> 13 Then Exit Sub
If Val(Grid_cons.TextMatrix(Grid_cons.Row, 0)) = 0 Then Exit Sub
PSMOV_LLAVE.rdoParameters(0) = LK_CODCIA
'PSMOV_LLAVE.rdoParameters(1) = Grid_cons.TextMatrix(Grid_cons.Row, 1)
PSMOV_LLAVE.rdoParameters(1) = Grid_cons.TextMatrix(Grid_cons.Row, 0)
PSMOV_LLAVE.rdoParameters(3) = LK_FECHA_COP1
PSMOV_LLAVE.rdoParameters(4) = LK_FECHA_COP1

mov_llave.Requery
If mov_llave.EOF Then
   MsgBox "Error de Datos..."
   Exit Sub
Else
   Grid_cons.Visible = False
   SSALIR.Visible = False
   ESTADO.Visible = True
   Frame3.Visible = True
   GoSub busca_fecha
   GoSub busca_voucher
   'Voucher_LostFocus
   'If mov_llave!MOV_MONEDA = "S" Then i_moneda.ListIndex = 0
   'If mov_llave!MOV_MONEDA = "D" Then i_moneda.ListIndex = 1
   GoSub busca_tipdoc
   i_numser.Text = mov_llave!MOV_serie
   i_numfac.Text = mov_llave!MOV_numfac
   i_glosa.Text = mov_llave!MOV_GLOSA
   loc_voucher = mov_llave!MOV_NRO_VOUCHER
   fila = 1
   Do Until mov_llave.EOF
      fila = fila + 1
      grid_fac.Rows = fila + 1
      grid_fac.TextMatrix(fila, 1) = Trim(mov_llave!MOV_CODCTA)
      grid_fac.TextMatrix(fila, 11) = mov_llave!MOV_DH
      SQ_OPER = 1
      PUB_CUENTA = mov_llave!MOV_CODCTA
      pu_codcia = LK_CODCIA
      PUB_CODCIA = LK_CODCIA
      LEER_COM_LLAVE
      If Not com_llave.EOF Then grid_fac.TextMatrix(fila, 2) = com_llave!com_DESCRIPCION
      If Val(com_llave!com_nivel) <> cop_llave!cop_nivel_max Then grid_fac.TextMatrix(fila, 2) = ""
      If mov_llave!MOV_DH = "D" Then
         grid_fac.TextMatrix(fila, 3) = mov_llave!MOV_IMPORTE
      Else
         grid_fac.TextMatrix(fila, 4) = mov_llave!MOV_IMPORTE
      End If
      If Nulo_Valor0(mov_llave!MOV_codclie) <> 0 Then
        SQ_OPER = 1
        pu_cp = Trim(mov_llave!MOV_CP)
        pu_codcia = LK_CODCIA
        pu_codclie = Val(mov_llave!MOV_codclie)
        LEER_CLI_LLAVE
        If Not cli_llave.EOF Then
           grid_fac.TextMatrix(fila, 5) = Nulo_Valor0(mov_llave!MOV_codclie)
           grid_fac.TextMatrix(fila, 6) = cli_llave!cli_nombre
           grid_fac.TextMatrix(fila, 7) = mov_llave!MOV_CP
           grid_fac.TextMatrix(fila, 9) = mov_llave!MOV_CP
        End If
        
      End If
      
      mov_llave.MoveNext
   Loop
End If

 suma_grid
 suma_subtotal

Exit Sub
   
   
   
   

busca_fecha:
i_fecha2.Text = Format(Grid_cons.TextMatrix(Grid_cons.Row, 1), "dd/mm/yyyy")

busca_tipdoc:
fila = 0
Do Until Val(i_tipdoc.Text) = Val(mov_llave!MOV_SUNAT) Or fila > 100
   i_tipdoc.ListIndex = fila
   fila = fila + 1
Loop
If fila > 100 Then MsgBox "Error de Fechas..."
Return

busca_voucher:
fila = 0
Do Until (voucher.ItemData(voucher.ListIndex)) = Val(mov_llave!MOV_TIPMOV) Or fila > 100
   voucher.ListIndex = fila
   fila = fila + 1
Loop
If fila > 100 Then MsgBox "Error de Voucher..."
Return



End Sub


Private Sub grid_docu_KeyPress(KeyAscii As Integer)
Dim XFILA As Integer
Dim wscadena  As String
Dim WS_anterior As Currency
Dim WS_ACTUAL As Currency
Dim wmensa As String
Dim wx_importe As Currency
If KeyAscii = 27 Then
   frmdocu.Visible = False
   grid_fac.SetFocus
End If
If KeyAscii <> 13 Then Exit Sub
   If grid_docu.Rows <= 1 Then
      frmdocu.Visible = False
      grid_fac.SetFocus
     Exit Sub
   End If
   frmdocu.Visible = False
   grid_fac.TextMatrix(grid_fac.Row, 8) = grid_docu.TextMatrix(grid_docu.Row, 3)
   'If Trim(i_moneda.Text) = "D" Then
   '  wx_importe = grid_docu.TextMatrix(grid_docu.Row, 8)
   ' Else
     wx_importe = grid_docu.TextMatrix(grid_docu.Row, 6)
   'End If
   
   'If grid_fac.TextMatrix(grid_fac.Row, 11) = "D" Then
   '   grid_fac.TextMatrix(grid_fac.Row, 3) = wx_importe
   '   grid_fac.Col = 3
   'Else
      grid_fac.TextMatrix(grid_fac.Row, 4) = wx_importe
      grid_fac.Col = 4
   'End If
   If voucher.ItemData(voucher.ListIndex) <> 3 Then GoTo saltar

'If Trim(grid_docu.TextMatrix(grid_docu.Row, 10)) = "D" Then
'Print Trim(grid_docu.TextMatrix(grid_docu.Row, 10))
If Trim(i_moneda.Text) = "" Then
   MsgBox "Ingrese que Moneda ", 48, Pub_Titulo
   Exit Sub
End If
'If Trim(i_moneda.Text) = "D" Then
'   WS_ACTUAL = Format(Val(grid_docu.TextMatrix(grid_docu.Row, 8)) * Val(tc.Text), "0.00")
'   WS_anterior = Format(Val(grid_docu.TextMatrix(grid_docu.Row, 6)), "0.00")
'Else
   WS_ACTUAL = Format(Val(grid_docu.TextMatrix(grid_docu.Row, 6)), "0.00")
   WS_anterior = Format(Val(grid_docu.TextMatrix(grid_docu.Row, 6)), "0.00")
'End If
   wmensa = "Monto Provision  S/. : " & WS_anterior & " (T.c.: " & Format(grid_docu.TextMatrix(grid_docu.Row, 9), "0.0000") & ")" & Chr(13) & "Monto Actual     S/. : " & WS_ACTUAL & " (T.c.: " & Format(tc.Text, "0.0000") & ")" & Chr(13) & Chr(13) & "Diferencia x T. de Cambio S/. : " & Format(WS_anterior - WS_ACTUAL, "0.00")
   FLAG_DIF_TC = ""
   LOC_DIF_TC = 0
   CUENTA_DIF = ""
   LOC_CODCLIE = 0
   loc_cp = ""
   LOC_DOCU = ""
   If (WS_anterior - WS_ACTUAL) <> 0 Then
       MsgBox wmensa, 48, Pub_Titulo
       FLAG_DIF_TC = "A"
       LOC_DIF_TC = WS_anterior - WS_ACTUAL
       CUENTA_DIF = grid_fac.TextMatrix(grid_fac.Row, 1)
       LOC_CODCLIE = grid_fac.TextMatrix(grid_fac.Row, 5)
       loc_cp = Trim(grid_fac.TextMatrix(grid_fac.Row, 4))
       LOC_DOCU = Trim(grid_fac.TextMatrix(grid_fac.Row, 8))
   End If
   men.Caption = "Dif. T.C. " & Format(WS_anterior - WS_ACTUAL, "0.00")
saltar:
   wscadena = FORM_CONTA2.grid_fac.TextMatrix(grid_fac.Row, 8)
   For XFILA = 0 To i_tipdoc.ListCount - 1
    i_tipdoc.ListIndex = XFILA
    If Left(i_tipdoc.Text, 2) = Left(wscadena, 2) Then
    Exit For
    End If
   Next
   i_numser.Text = Val(Mid(wscadena, 4, 6))
   i_numfac.Text = Val(Mid(wscadena, 8, 16))

   grid_fac.SetFocus
   Azul3 TEXTOVAR, TEXTOVAR

End Sub





Private Sub liscc_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   liscc.Visible = False
   grid_fac.SetFocus
End If
If KeyAscii = 13 Then
   TEXTOVAR2.Visible = True
   If TEXTOVAR2.Visible = True Then
       TEXTOVAR2.Text = Trim(Right(liscc.Text, 8))
       liscc.Visible = False
       TEXTOVAR2.SetFocus
   End If
End If

End Sub

Private Sub liscc_LostFocus()
liscc.Visible = False
End Sub



Private Sub textovar_Change()
On Error GoTo verlod
PA:
If grid_fac.Col = 0 Then
   If IsDate(TEXTOVAR.Text) Then
   '   grid_fac.Text = Format(TEXTOVAR.Text, "dd/mm/yyyy")
      Exit Sub
   End If
End If

'If grid_fac.COL = 0 Then
  grid_fac.Text = TEXTOVAR.Text
'ElseIf grid_fac.COL = 8 Then
 ' grid_fac.Text = Trim(TEXTOVAR.Text)
'Else
'  grid_fac.Text = Format(TEXTOVAR.Text, "0.00")
'End If
If grid_fac.Col = 11 Or grid_fac.Col = 12 Then
 suma_grid
 
End If
Exit Sub
verlod:
Resume Next
Exit Sub
End Sub

Private Sub TEXTOVAR_GotFocus()
temporal = grid_fac.TextMatrix(grid_fac.Row, grid_fac.Col)
Azul TEXTOVAR, TEXTOVAR
End Sub

Private Sub textovar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Exit Sub

If grid_fac.Col = 5 Then
  If LV_CLI.Visible = True Then GoTo SALTAX
End If

 suma_grid
 
'If F2 = 1 And KeyCode <> 13 Then Exit Sub

If KeyCode = 113 Then
   F2 = 1
   Exit Sub
End If
If KeyCode = 37 Or KeyCode = 38 Or KeyCode = 39 Or KeyCode = 13 Or KeyCode = 40 Then
Else
Exit Sub
End If
If KeyCode = 13 Then
If grid_fac.TextMatrix(grid_fac.Row, 3) <> "" And grid_fac.TextMatrix(grid_fac.Row, 4) <> "" Then
   MsgBox "Solo Debe o Haber"
   If grid_fac.Col = 3 Then grid_fac.TextMatrix(grid_fac.Row, 4) = ""
   If grid_fac.Col = 4 Then grid_fac.TextMatrix(grid_fac.Row, 3) = ""
End If
End If



If KeyCode = 38 Then
   If grid_fac.Row = 2 Then Exit Sub
End If
   
If KeyCode = 37 Then
   If grid_fac.Col = 0 Then Exit Sub
End If
If KeyCode = 40 Then
   If grid_fac.Row = grid_fac.Rows - 1 Then Exit Sub
End If
If KeyCode = 13 And grid_fac.Row = grid_fac.Rows - 1 Then
      grid_fac.Rows = grid_fac.Rows + 1
      grid_fac.RowHeight(grid_fac.Rows - 1) = 285
      grid_fac.TextMatrix(grid_fac.Rows - 1, 10) = 1
      grid_fac.MergeRow(grid_fac.Rows - 1) = False
      grid_fac.Col = 1
      grid_fac.Row = grid_fac.Rows - 1
      grid_fac.SetFocus
      Exit Sub
End If
   



If KeyCode = 38 Then
   grid_fac.Row = grid_fac.Row - 1
   FLAG_SALIR = 9
   Exit Sub
ElseIf KeyCode = 40 Then
   FLAG_SALIR = 9
   grid_fac.Row = grid_fac.Row + 1
   Exit Sub
ElseIf KeyCode = 13 Then
   If Trim(grid_fac.TextMatrix(grid_fac.Row, 4)) = "C" Or Trim(grid_fac.TextMatrix(grid_fac.Row, 4)) = "P" Then
      grid_fac.Col = 5
      Exit Sub
   Else
      grid_fac.Row = grid_fac.Row + 1
       Exit Sub
   End If
End If


If KeyCode = 37 Then
   grid_fac.Col = grid_fac.Col - 1
ElseIf KeyCode = 39 Then
   'If grid_fac.Col = 4 Then
   '   grid_fac.Col = 1
   'Else
      grid_fac.Col = grid_fac.Col + 1
   'End If
End If
Exit Sub
    
    
SALTAX:
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And TEXTOVAR.Text = "" Then
  loc_key = 1
  Set LV_CLI.SelectedItem = LV_CLI.ListItems(loc_key)
  LV_CLI.ListItems.Item(loc_key).Selected = True
  LV_CLI.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > LV_CLI.ListItems.Count Then loc_key = LV_CLI.ListItems.Count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > LV_CLI.ListItems.Count Then loc_key = LV_CLI.ListItems.Count
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
  LV_CLI.ListItems.Item(loc_key).Selected = True
  LV_CLI.ListItems.Item(loc_key).EnsureVisible
  TEXTOVAR.Text = Trim(LV_CLI.ListItems.Item(loc_key).Text) & " "
  DoEvents
  TEXTOVAR.SelStart = Len(TEXTOVAR.Text)
  DoEvents
fin:
    
    
End Sub
Private Sub textovar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  If Trim(TEXTOVAR.Text) = "" Then
     TEXTOVAR.Text = ""
  End If
  TEXTOVAR.Visible = False
  grid_fac.SetFocus
  LV_CLI.Visible = False
  Exit Sub
End If
If grid_fac.Col = 5 Then GoTo SALT

If grid_fac.Col = 3 Or grid_fac.Col = 4 Then Consistencias grid_fac, TEXTOVAR, KeyAscii
If KeyAscii <> 13 Then Exit Sub
cmdIngreso.Visible = True

If grid_fac.TextMatrix(grid_fac.Row, 3) <> "" And grid_fac.TextMatrix(grid_fac.Row, 4) <> "" Then
   MsgBox "Solo Debe o Haber", 48, Pub_Titulo
   If grid_fac.Col = 3 Then grid_fac.TextMatrix(grid_fac.Row, 4) = ""
   If grid_fac.Col = 4 Then grid_fac.TextMatrix(grid_fac.Row, 3) = ""
   suma_grid
   
End If

If Trim(grid_fac.TextMatrix(grid_fac.Row, 4)) = "C" Or Trim(grid_fac.TextMatrix(grid_fac.Row, 4)) = "P" Then
      grid_fac.Col = 5
      'grid_fac.SetFocus
      'TEXTOVAR.SetFocus
      'grid_fac_KeyPress 13
      Exit Sub
End If

If grid_fac.Row = grid_fac.Rows - 1 Then
      grid_fac.Rows = grid_fac.Rows + 1
      grid_fac.TextMatrix(grid_fac.Rows - 1, 10) = 1
      grid_fac.RowHeight(grid_fac.Rows - 1) = 285
      grid_fac.MergeRow(grid_fac.Rows - 1) = False
      grid_fac.Col = 1
      TEXTOVAR.Visible = False
      grid_fac.Row = grid_fac.Rows - 1
      grid_fac.SetFocus
Else
        TEXTOVAR.Visible = False
        grid_fac.Row = grid_fac.Row + 1
        If grid_fac.Col = 3 Or grid_fac.Col = 4 Then
          If Trim(grid_fac.TextMatrix(grid_fac.Row, 4)) = "P" Or Trim(grid_fac.TextMatrix(grid_fac.Row, 4)) = "C" Then
                grid_fac.Col = 5
          End If
        Else
          If Trim(grid_fac.TextMatrix(grid_fac.Row, 11)) = "D" Then
              grid_fac.Col = 3
          Else
              grid_fac.Col = 4
          End If
       End If
        Exit Sub

    '  grid_fac.SetFocus
End If

Exit Sub


SALT:
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As ListItem    ' Variable FoundItem.
If KeyAscii = 27 Then
 TEXTOVAR2.Text = ""
 Exit Sub
End If

If KeyAscii <> 13 Then
   If grid_fac.Col = 5 Then grid_fac.TextMatrix(grid_fac.Row, 6) = ""
   GoTo fin
End If
F2 = 0
On Error GoTo OJO
pu_codclie = Val(TEXTOVAR.Text)
On Error GoTo 0
If Len(TEXTOVAR.Text) = 0 Then
   Exit Sub
End If
If pu_codclie <> 0 And IsNumeric(TEXTOVAR.Text) = True Then
   On Error GoTo OJO
   SQ_OPER = 1
   pu_cp = Trim(grid_fac.TextMatrix(grid_fac.Row, 4))
   pu_codcia = LK_CODCIA
   PUB_CODCLIE = TEXTOVAR.Text
   LEER_CLI_LLAVE
   On Error GoTo 0
   If cli_llave.EOF Then
    Azul TEXTOVAR, TEXTOVAR
    MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
    TEXTOVAR.SetFocus
    GoTo fin
   Else
        grid_fac.TextMatrix(grid_fac.Row, 6) = Trim(cli_llave!cli_nombre)
        grid_fac.TextMatrix(grid_fac.Row, 7) = Trim(cli_llave!CLI_CP)
        If 0 = 1 Or 0 = 2 Then
          If (grid_fac.Rows - 1) = grid_fac.Row Then
           TEXTOVAR.Visible = False
           grid_fac.SetFocus
           Exit Sub
          End If
          grid_fac.Row = grid_fac.Row + 1
          If Trim(grid_fac.TextMatrix(grid_fac.Row, 11)) = "D" Then
              grid_fac.Col = 3
          Else
              grid_fac.Col = 4
          End If
          
        Else
          grid_fac.Col = 8
          TEXTOVAR.Visible = False
          grid_fac.SetFocus
        End If
        Exit Sub
   End If
Else
On Error GoTo SIGUE

   If loc_key <> 0 Then valor = UCase(LV_CLI.ListItems.Item(loc_key).Text)
   If Trim(UCase(TEXTOVAR.Text)) = Left(valor, Len(Trim(TEXTOVAR.Text))) Then
   Else
      Exit Sub
   End If
   If loc_key = 0 Then Exit Sub
   TEXTOVAR.Text = Trim(LV_CLI.ListItems.Item(loc_key).SubItems(1))
   pu_cp = grid_fac.TextMatrix(grid_fac.Row, 4)
   pu_codcia = LK_CODCIA
   pu_codclie = TEXTOVAR.Text
   LEER_CLI_LLAVE
End If

LV_CLI.Visible = False
grid_fac.TextMatrix(grid_fac.Row, 6) = Trim(cli_llave!cli_nombre)
grid_fac.TextMatrix(grid_fac.Row, 7) = Trim(cli_llave!CLI_CP)
TEXTOVAR.Visible = False
If voucher.ItemData(voucher.ListIndex) = 1 Or voucher.ItemData(voucher.ListIndex) = 2 Then
  If (grid_fac.Rows - 1) = grid_fac.Row Then
   TEXTOVAR.Visible = False
   grid_fac.SetFocus
   Exit Sub
  End If
  grid_fac.Row = grid_fac.Row + 1
 
  If Trim(grid_fac.TextMatrix(grid_fac.Row, 11)) = "D" Then
      grid_fac.Col = 3
  Else
      grid_fac.Col = 4
  End If
Else
  grid_fac.Col = 8
  TEXTOVAR.Visible = False
  grid_fac.SetFocus
End If

Exit Sub

If grid_fac.Row = grid_fac.Rows - 1 Then
      grid_fac.Rows = grid_fac.Rows + 1
      grid_fac.TextMatrix(grid_fac.Rows - 1, 10) = 1
      grid_fac.RowHeight(grid_fac.Rows - 1) = 285
      grid_fac.MergeRow(grid_fac.Rows - 1) = False
      grid_fac.Col = 1
      grid_fac.Row = grid_fac.Rows - 1
      grid_fac.SetFocus
Else
      grid_fac.Row = grid_fac.Row + 1
End If
If grid_fac.TextMatrix(grid_fac.Row, 1) <> "" Then grid_fac.Col = 3

'TEXTOVAR.Visible = False
'grid_fac.SetFocus

Exit Sub
SIGUE:
If Err.Number = 35600 Then
  Exit Sub
End If
fin:
OJO:





End Sub

Private Sub textovar_KeyUp(KeyCode As Integer, Shift As Integer)

If grid_fac.Col <> 5 Then Exit Sub
Dim var, VAR2
'If Len(txt_key.Text) = 0 Or IsNumeric(txt_key.Text) = True Then
'   ListView1.Visible = False
'   Exit Sub
'End If

If Len(TEXTOVAR.Text) = 0 Or IsNumeric(TEXTOVAR.Text) Then
   LV_CLI.Visible = False
   Exit Sub
End If
If LV_CLI.Visible = False Or Len(Trim(TEXTOVAR.Text)) = 1 Then
   XX_CUENTA = Trim(grid_fac.TextMatrix(grid_fac.Row, 4))
   loc_key = 0
    var = Asc(TEXTOVAR.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    VAR2 = Val(XX_CUENTA) + 1
    
    If Trim(XX_CUENTA) <> "" Then
       PUB_CP = XX_CUENTA
       numarchi = 1
       archi = "SELECT CLI_CODCLIE , CLI_CODCIA, CLI_CP, CLI_NOMBRE, CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM, CLI_RUC_ESPOSO   FROM CLIENTES WHERE CLI_CP = '" & PUB_CP & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY CLI_NOMBRE"
    End If
       
    PROC_LISVIEW LV_CLI
    loc_key = 0
    If LV_CLI.Visible Then
     loc_key = 1
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If LV_CLI.Visible Then
  Set itmFound = LV_CLI.FindItem(LTrim(TEXTOVAR.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = Val(itmFound.Tag)
   If loc_key + 8 > LV_CLI.ListItems.Count Then
      LV_CLI.ListItems.Item(LV_CLI.ListItems.Count).EnsureVisible
   Else
     LV_CLI.ListItems.Item(loc_key + 8).EnsureVisible
   End If
  End If
  Exit Sub
End If

End Sub

Private Sub grid_fac_DblClick()
Dim suma
    For fila = 2 To grid_fac.Rows - 1
        If Left(grid_fac.TextMatrix(fila, 1), 2) = "40" Then
            suma = suma + Val(grid_fac.TextMatrix(fila, 4))
        End If
    Next fila
'MsgBox suma
    cmdIngreso.Visible = True
End Sub

Private Sub grid_fac_EnterCell()
'If Trim(i_plantilla.Text) = "" Then Exit Sub
lblinforme.Caption = ""
If grid_fac.Col = 3 Then
    lblinforme.Caption = Trim(grid_fac.TextMatrix(grid_fac.Row, 19))
    If Trim(grid_fac.TextMatrix(grid_fac.Row, 4)) = "" Then
       grid_fac.TextMatrix(1, 3) = ""
    ElseIf Trim(grid_fac.TextMatrix(grid_fac.Row, 4)) = "C" Then
       grid_fac.TextMatrix(1, 3) = "Cliente"
    ElseIf Trim(grid_fac.TextMatrix(grid_fac.Row, 4)) = "P" Then
       grid_fac.TextMatrix(1, 3) = "Proveedor"
    ElseIf Trim(grid_fac.TextMatrix(grid_fac.Row, 4)) = "E" Then
       grid_fac.TextMatrix(1, 3) = "Empleado"
    End If
End If
If grid_fac.Col = 6 Then lblinforme.Caption = Trim(grid_fac.TextMatrix(grid_fac.Row, 21))
If grid_fac.Col = 7 Then lblinforme.Caption = Trim(grid_fac.TextMatrix(grid_fac.Row, 22))
If grid_fac.Col = 5 Then lblinforme.Caption = Trim(grid_fac.TextMatrix(grid_fac.Row, 23))

TEXTOVAR.Visible = False
TEXTOVAR2.Visible = False
I_NUMFAC2.Visible = False
grid_fac.SetFocus
If grid_fac.Col = 2 Then
   TEXTOVAR2.Visible = False
   Exit Sub
End If
If grid_fac.Col >= 17 Then
   TEXTOVAR2.Visible = False
   Exit Sub
End If

TEXTOVAR.Visible = False
TEXTOVAR2.Visible = False
I_NUMFAC2.Visible = False
grid_fac.SetFocus
If grid_fac.Text <> "" And (grid_fac.Col = 1) Then Exit Sub
If grid_fac.TextMatrix(grid_fac.Row, 1) = "" Then Exit Sub

SSSS:
   If grid_fac.CellWidth <= 0 Then Exit Sub
   Exit Sub
    If grid_fac.Col = 1 Then
     XX_CUENTA = Trim(grid_fac.TextMatrix(grid_fac.Row, 4))
     TEXTOVAR2.Left = ESTADO.Left + grid_fac.Left + grid_fac.CellLeft
     TEXTOVAR2.Width = grid_fac.CellWidth
     TEXTOVAR2.Height = grid_fac.CellHeight
     TEXTOVAR2.Top = grid_fac.Top + grid_fac.CellTop   '- 340 '480
     TEXTOVAR2.Text = Trim(grid_fac.TextMatrix(grid_fac.Row, grid_fac.Col))
 
     wfila_act = grid_fac.Row
     TEXTOVAR2.Visible = True
     TEXTOVAR.Visible = False
     Azul TEXTOVAR2, TEXTOVAR2
    Else
     TEXTOVAR.Left = ESTADO.Left + grid_fac.Left + grid_fac.CellLeft
     TEXTOVAR.Width = grid_fac.CellWidth
     TEXTOVAR.Height = grid_fac.CellHeight
     TEXTOVAR.Top = ESTADO.Top + grid_fac.Top + grid_fac.CellTop - 1920  '480
     TEXTOVAR.Text = Trim(grid_fac.TextMatrix(grid_fac.Row, grid_fac.Col))
     If grid_fac.Col = 5 And Trim(grid_fac.TextMatrix(grid_fac.Row, 9)) = "" Then Exit Sub
     wfila_act = grid_fac.Row
    End If

End Sub


Private Sub grid_fac_KeyPress(KeyAscii As Integer)
'If Trim(i_plantilla.Text) = "" Then Exit Sub
Dim WTIPMOV  As String
Dim WS_SALDO As Currency
Dim wmes
Dim wfecha
Dim wnrovoucher
Dim WS_TIPO_CAMBIO As Currency
Dim WWFBG  As Currency
Dim wwnumfac  As Currency
Dim WWSERIE As Currency
Dim WDEBE As Currency
Dim WHABER As Currency
Dim a As Long
Dim t, WC
Static CONS
If KeyAscii = 27 Then
  If grid_fac.Col = 8 Then
   grid_fac.TextMatrix(grid_fac.Row, 8) = ""
  End If
  Exit Sub
End If
'If KeyAscii <> 13 Then Exit Sub
If grid_fac.Col = 0 Then Exit Sub
If grid_fac.Col = 2 Then Exit Sub
'If grid_fac.COL = 12 Then Exit Sub
'If grid_fac.COL = 2 Then Exit Sub
If grid_fac.Col = 3 And Trim(grid_fac.TextMatrix(grid_fac.Row, 4)) = "" Then '9898*********************
    pub_mensaje = InputBox("Si desea ingresar un Cliente Digite la letra C , si desea Proveedor la P ", "Ingresar Cliente /Proveedor", "")
    If pub_mensaje = "" Then
       grid_fac.SetFocus
       Exit Sub
    End If
    
    If UCase(pub_mensaje) = "C" Or UCase(pub_mensaje) = "P" Then
      grid_fac.TextMatrix(grid_fac.Row, 4) = UCase(pub_mensaje)
      grid_fac.SetFocus
      Exit Sub
    Else
       Exit Sub
    End If
End If

'If grid_fac.Col = 5 And Trim(grid_fac.TextMatrix(grid_fac.Row, 9)) = "" Then Exit Sub


    If grid_fac.Col = 3 Or grid_fac.Col = 4 Then
       'If grid_fac.TextMatrix(grid_fac.Row, 11) = "D" And grid_fac.Col = 4 Then Exit Sub
       'If grid_fac.TextMatrix(grid_fac.Row, 11) = "H" And grid_fac.Col = 3 Then Exit Sub
    End If




If grid_fac.Col = 1 Then
   a = Val(grid_fac.TextMatrix(grid_fac.Row - 1, 0))
   a = a + 1
  grid_fac.TextMatrix(grid_fac.Row, 0) = a
End If

If grid_fac.Col = 3 Then
  If Val(grid_fac.TextMatrix(grid_fac.Row, 1)) = 0 Then Exit Sub
End If

If grid_fac.Col = 4 Then
  Exit Sub
End If
    If grid_fac.Col = 3 Or grid_fac.Col = 4 Then
       If grid_fac.TextMatrix(grid_fac.Row, 11) = "D" Then grid_fac.Col = 3
       If grid_fac.TextMatrix(grid_fac.Row, 11) = "H" Then grid_fac.Col = 4
    End If
    TEXTOVAR.Visible = False
    TEXTOVAR2.Visible = False
    I_NUMFAC2.Visible = False
    If grid_fac.Col <> 7 Then
        XX_CUENTA = Trim(grid_fac.TextMatrix(grid_fac.Row, 4))
        TEXTOVAR2.Left = ESTADO.Left + grid_fac.Left + grid_fac.CellLeft
        TEXTOVAR2.Width = grid_fac.CellWidth
        TEXTOVAR2.Height = grid_fac.CellHeight
        TEXTOVAR2.Top = grid_fac.CellTop + grid_fac.Top  ' TOP'+ grid_fac.CellTop '- 340 '480
        TEXTOVAR2.Text = grid_fac.TextMatrix(grid_fac.Row, grid_fac.Col)
        wfila_act = grid_fac.Row
        TEXTOVAR2.Visible = True
        Azul TEXTOVAR2, TEXTOVAR2
        TEXTOVAR2.SetFocus
    ElseIf grid_fac.Col = 7 Then
        cboSunat.Left = ESTADO.Left + grid_fac.Left + grid_fac.CellLeft
        cboSunat.Width = 1200
        cboSunat.Top = grid_fac.CellTop + grid_fac.Top  ' TOP'+ grid_fac.CellTop '- 340 '480
        wfila_act = grid_fac.Row
        cboSunat.Visible = True
        cboSunat.SetFocus
        Res = SendMessageLong(cboSunat.hwnd, &H14F, True, 0)
    '================bloqueado por kmi ya que nunca entra
'    ElseIf grid_fac.Col = 893 And Val(grid_fac.TextMatrix(grid_fac.Row, 5)) <> 0 Then
'      ' GoSub LLENA_DATOS
'       'I_NUMFAC2.Left = ESTADO.Left + grid_fac.Left + grid_fac.CellLeft
'       'I_NUMFAC2.Width = grid_fac.CellWidth
'       'I_NUMFAC2.Top = ESTADO.Top + grid_fac.Top + grid_fac.CellTop - 2040
'       'I_NUMFAC2.Visible = True
'       'I_NUMFAC2.SetFocus
'    '   SendKeys "%{up}"
'     '  ElseIf grid_fac.Col <> 8 Then
'        If grid_fac.CellWidth < 0 Then Exit Sub
'        TEXTOVAR.Left = ESTADO.Left + grid_fac.Left + grid_fac.CellLeft
'        TEXTOVAR.Width = grid_fac.CellWidth
'        TEXTOVAR.Height = grid_fac.CellHeight
'        TEXTOVAR.Top = ESTADO.Top + grid_fac.Top + grid_fac.CellTop - 2050 '480
'        TEXTOVAR.Text = grid_fac.TextMatrix(grid_fac.Row, grid_fac.Col)
'        wfila_act = grid_fac.Row
'        TEXTOVAR.Visible = True
'        Azul TEXTOVAR, TEXTOVAR
    End If
    If KeyAscii <> 13 Then
     TEXTOVAR2.Text = Chr(KeyAscii)
     TEXTOVAR2.SelStart = Len((TEXTOVAR2.Text))
    End If

Exit Sub
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
Exit Sub


End Sub

Private Sub grid_fac_KeyUp(KeyCode As Integer, Shift As Integer)
Dim IMP_MONEDA As String
Dim IMP_INI As Currency
Dim IMP_TC As Currency
Dim WTIPMOV  As String
Dim WS_SALDO As Currency
Dim wmes
Dim wfecha
Dim wnrovoucher
Dim WS_TIPO_CAMBIO As Currency
Dim WWFBG  As Currency
Dim wwnumfac  As Currency
Dim WWSERIE As Currency
Dim WDEBE As Currency
Dim WHABER As Currency

Dim WC
Dim a, WF As Integer
Dim tf, t, tc
Dim SALE As Boolean
If KeyCode = 45 And grid_fac.Col = 8 Then
    GoSub LLENA_DATOS
    Exit Sub
End If


If KeyCode = 113 And grid_fac.Col = 12 Then
  WC = InputBox("Fecha Para Capturar el Tipo de Cambio : ", "Fecha", 1)
  If WC = "" Then Exit Sub
  If Not IsDate(WC) Then
    MsgBox "Fecha no procede.", 48, Pub_Titulo
    grid_fac.SetFocus
    Exit Sub
  End If
  PUB_CAL_INI = Format(WC, "dd/mm/yyyy")
  PUB_CAL_FIN = Format(WC, "dd/mm/yyyy")
  PUB_CODCIA = LK_CODCIA
  LEER_CAL_LLAVE
  If cal_llave.EOF Then
    MsgBox "No tiene T.C. en esta fecha Verificar 0 Ingresar T.C. ", 48, Pub_Titulo
    Exit Sub
  End If
  If Val(Format(cal_llave!cal_tipo_cambio, "0.0000")) = 0 Then
    MsgBox "No tiene T.C. en esta fecha Verificar 0 Ingresar T.C. ", 48, Pub_Titulo
    Exit Sub
  End If
  grid_fac.TextMatrix(grid_fac.Row, 24) = Format(cal_llave!cal_tipo_cambio, "0.0000")
  grid_fac.TextMatrix(grid_fac.Row, 25) = "A"
  grid_fac.SetFocus
  
End If
If KeyCode = 113 And grid_fac.Col = 3 Then ' PUB_CP
    lblinforme.Caption = ""
    grid_fac.TextMatrix(grid_fac.Row, 3) = ""
    If Trim(grid_fac.TextMatrix(grid_fac.Row, 4)) = "" Then
       grid_fac.TextMatrix(grid_fac.Row, 4) = "C"
       grid_fac.TextMatrix(1, 3) = "Cliente"
    ElseIf Trim(grid_fac.TextMatrix(grid_fac.Row, 4)) = "C" Then
       grid_fac.TextMatrix(grid_fac.Row, 4) = "P"
       grid_fac.TextMatrix(1, 3) = "Proveedor"
    ElseIf Trim(grid_fac.TextMatrix(grid_fac.Row, 4)) = "P" Then
       grid_fac.TextMatrix(grid_fac.Row, 4) = "E"
       grid_fac.TextMatrix(1, 3) = "Empleado"
    ElseIf Trim(grid_fac.TextMatrix(grid_fac.Row, 4)) = "E" Then
       grid_fac.TextMatrix(grid_fac.Row, 4) = " "
       grid_fac.TextMatrix(1, 3) = ""
    End If
    Exit Sub
End If

If KeyCode = 113 And grid_fac.Col = 6 Then ' Centros de Costos.
    PUB_TIPREG = 40
    PUB_CODCIA = LK_CODCIA
    Load FrmDatArti
    FrmDatArti.Caption = "Codigo de Centro de Costo " & " TAB_TIPREG = " & PUB_TIPREG
    FrmDatArti.Show 1
    Exit Sub
End If
If KeyCode = 113 And grid_fac.Col = 5 Then ' Codigo de Sunat
    PUB_TIPREG = 240
    PUB_CODCIA = LK_CODCIA
    Load FrmDatArti
    FrmDatArti.Caption = "Codigo OPC " & " TAB_TIPREG = " & PUB_TIPREG
    FrmDatArti.Show 1
    Exit Sub
End If
If KeyCode = 113 And grid_fac.Col = 7 Then ' Codigo de Sunat
    PUB_TIPREG = 50
    PUB_CODCIA = "00"
    Load FrmDatArti
    FrmDatArti.Caption = "Codigo SUNAT " & " TAB_TIPREG = " & PUB_TIPREG
    FrmDatArti.Show 1
    Exit Sub
End If
If KeyCode = 115 Then  ' Codigo de Sunat
   tc = grid_fac.Row
   grid_fac.AddItem ""
   For a = 0 To grid_fac.Cols - 1
     grid_fac.TextMatrix(grid_fac.Rows - 1, a) = grid_fac.TextMatrix(tc, a)
   Next a
   Exit Sub
End If

If KeyCode = 45 Then
   If grid_fac.Row = (grid_fac.Rows - 1) Then
     grid_fac.Rows = grid_fac.Rows + 1
     grid_fac.RowHeight(grid_fac.Rows - 1) = 285
     grid_fac.TextMatrix(grid_fac.Rows - 1, 10) = 1
     grid_fac.Row = grid_fac.Rows - 1
     grid_fac.Col = 1
     grid_fac.SetFocus
   Else
     grid_fac.AddItem " ", grid_fac.Row
     grid_fac.TextMatrix(grid_fac.Rows - 1, 10) = 1
     grid_fac.SetFocus
     grid_fac.Col = 1
     grid_fac.RowHeight(grid_fac.Row) = 285
   End If
End If

If cop_llave!COP_FLAG_MAYORIZACION = "M" Then
 'MsgBox "Ojo estaba Mayorizado..."
End If
 If KeyCode = 46 Then
   If grid_fac.Row = 2 And grid_fac.Rows <> 3 Then
     grid_fac.RemoveItem grid_fac.Row
     suma_grid
     Exit Sub
   End If
   If grid_fac.Row > 2 Then grid_fac.RemoveItem grid_fac.Row
   suma_grid
   Exit Sub
 End If

'If Left(grid_fac.TextMatrix(grid_fac.Row, 0), 2) <> "MA" Then Exit Sub
 If KeyCode = 32 Then
   Exit Sub
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
   Exit Sub
    If grid_fac.Row >= grid_fac.Rows - 1 Then Exit Sub
    Wsec = Wsec + 1
    If Trim(grid_fac.TextMatrix(grid_fac.Row + 1, 11)) = "8" Then
         Exit Sub
    Else
      If Trim(grid_fac.TextMatrix(grid_fac.Row + 1, 0)) = "T" Then Exit Sub
    End If
    If Val(grid_fac.TextMatrix(grid_fac.Row, 4)) = 0 And Val(grid_fac.TextMatrix(grid_fac.Row, 5)) = 0 Then Exit Sub
    grid_fac.AddItem "", grid_fac.Row + 1
    grid_fac.TextMatrix(grid_fac.Rows - 1, 10) = 1
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
Exit Sub
If KeyCode = 46 Then
If grid_fac.Rows <= 3 Then
Else
   pub_mensaje = MsgBox("Desea Quitar el Item de la Cuenta : " & Trim(grid_fac.TextMatrix(grid_fac.Row, 1)), vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
   If pub_mensaje = vbNo Then
     grid_fac.SetFocus
     Exit Sub
   Else
    '   grid_fac.RowHeight(grid_fac.Row) = 1
    '   grid_fac.Row = grid_fac.Row + 1
       grid_fac.RemoveItem (grid_fac.Row)
       grid_fac.Refresh
       suma_grid
       grid_fac.SetFocus
   End If
End If
End If
'grid_fac.SetFocus
Exit Sub

LLENA_DATOS:


Return


End Sub

Private Sub i_fecha2_Change()
 tc.Text = ""
End Sub

Private Sub i_fecha2_GotFocus()
If grid_fac.Visible = False Then Exit Sub
Azul2 i_fecha2, i_fecha2
'CABE_MAN

End Sub

Private Sub i_fecha2_KeyPress(KeyAscii As Integer)
'On Error GoTo pasa
Dim Wflag As String * 1
Dim wsFECHA1, WS_FECHA2
If KeyAscii <> 13 Then Exit Sub
txtdetalle.SetFocus


End Sub

Private Sub i_fecha2_LostFocus()
If Right(i_fecha2.Text, 2) = "__" Then
     i_fecha2.Text = Format(Left(i_fecha2.Text, 8), "dd/mm/yyyy")
End If
If Not IsDate(i_fecha2.Text) Then
  MsgBox "Fecha no procede.", 48, Pub_Titulo
  i_fecha2.SetFocus
  Exit Sub
End If
End Sub

Private Sub i_glosa_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
i_fecha2.SetFocus

End Sub

Private Sub i_moneda_Click()
On Error GoTo SALE
Dim wtc
If grid_fac.Cols = 2 Then Exit Sub
tc.Text = ""
If Trim(i_moneda.Text) = "D" Then
    SQ_OPER = 1
    PUB_CAL_INI = i_fecha2.Text
    PUB_CAL_FIN = i_fecha2.Text
    PUB_CODCIA = LK_CODCIA
    LEER_CAL_LLAVE
    If cal_llave.EOF Then
      MsgBox "Verificar Fecha para tipo de Cambio.", 48, Pub_Titulo
      Exit Sub
    End If
    If Val(Format(cal_llave!cal_tipo_cambio, "0.0000")) = 0 Then
       wtc = InputBox("Fecha no Tiene T.C.  Ingrese : ")
       If wtc = "" Then Exit Sub
       If Val(wtc) = 0 Then
          MsgBox "No procede.", 48, Pub_Titulo
          Exit Sub
       End If
       cal_llave.Edit
       cal_llave!cal_tipo_cambio = Val(wtc)
       cal_llave.Update
    End If
   tc.Text = Format(cal_llave!cal_tipo_cambio, "0.0000")
   If Option1(1) = False Then
   For fila = 2 To grid_fac.Rows - 1
         grid_fac.TextMatrix(fila, 24) = tc.Text
   Next fila
   End If
Else
   For fila = 2 To grid_fac.Rows - 1
         grid_fac.TextMatrix(fila, 24) = tc.Text
         grid_fac.TextMatrix(fila, 25) = ""
   Next fila
End If
Exit Sub
SALE:
 MsgBox Err.Description & " Intente Nuevamente", 48, Pub_Titulo

End Sub

Private Sub i_moneda_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = 45 Then
'    Load frmTC
'    frmTC.Show 1
'End If

End Sub

Private Sub i_numfac_GotFocus()
Azul i_numfac, i_numfac
End Sub

Private Sub i_numfac_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii <> 13 Then Exit Sub
grid_fac.Col = 1
txtdetalle.SetFocus
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

Private Sub i_numser_GotFocus()
Azul i_numser, i_numser
End Sub

Private Sub i_numser_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii <> 13 Then Exit Sub
i_numfac.SetFocus


End Sub

Private Sub i_plantilla_KeyPress(KeyAscii As Integer)
Dim IdVoucher As Integer
If voucher.ListIndex >= 0 Then
    IdVoucher = voucher.ItemData(voucher.ListIndex)
Else
    Exit Sub
End If
If KeyAscii <> 13 Then Exit Sub

TEXTOVAR.Visible = False
TEXTOVAR2.Visible = False
cancelar_Click

'ESTADO.Caption = "Estado :   < Ingreso de VOUCHER >"
FORM_CONTA2.i_glosa.Enabled = True
frmcon.Visible = False
Option1(0).Value = True
WPASA = True
CABE_ING

fila = 1
SUM_D = 0
SUM_H = 0

PSTEMP_MAYOR.rdoParameters(0) = LK_CODCIA
PSTEMP_MAYOR.rdoParameters(1) = IdVoucher
PSTEMP_MAYOR.rdoParameters(2) = Val(Right(i_plantilla.Text, 5))
grid_fac.TextMatrix(2, 4) = Trim(Right(voucher.Text, 5))
temp_mayor.Requery
Do Until temp_mayor.EOF

   If temp_mayor!PLT_NUMERO <> Val(Right(i_plantilla.Text, 5)) Then Exit Do
   
   If temp_mayor!PLT_SECUENCIA > 0 Then
   
      SQ_OPER = 1
      PUB_CUENTA = temp_mayor!PLT_CUENTA
      pu_codcia = LK_CODCIA
      PUB_CODCIA = LK_CODCIA
      LEER_COM_LLAVE
      If com_llave.EOF Then GoTo P_P
        fila = fila + 1
        grid_fac.Rows = fila + 1
        grid_fac.RowHeight(grid_fac.Rows - 1) = 285
        grid_fac.TextMatrix(fila, 0) = Format(fila - 1, "00")
        grid_fac.TextMatrix(fila, 1) = Trim(temp_mayor!PLT_CUENTA)
        grid_fac.TextMatrix(fila, 4) = Trim(Nulo_Valors(temp_mayor!PLT_NOMBRE))
      If grid_fac.TextMatrix(fila, 4) = "C" Then
         grid_fac.TextMatrix(1, 3) = "Cliente"
      ElseIf grid_fac.TextMatrix(fila, 4) = "P" Then
         grid_fac.TextMatrix(1, 3) = "Proveedor"
      ElseIf grid_fac.TextMatrix(fila, 4) = "E" Then
         grid_fac.TextMatrix(1, 3) = "Personal"
      End If
      grid_fac.TextMatrix(fila, 10) = temp_mayor!PLT_CUENTA
      grid_fac.TextMatrix(fila, 18) = Trim(temp_mayor!PLT_DH)
      grid_fac.TextMatrix(fila, 10) = "1"
      i_glosa.Text = Trim(temp_mayor!PLT_GLOSA)
      If Not com_llave.EOF Then
        grid_fac.TextMatrix(fila, 2) = com_llave!com_DESCRIPCION
        If Val(com_llave!com_nivel) <> cop_llave!cop_nivel_max Then grid_fac.TextMatrix(fila, 2) = ""
      End If
      cmdIngreso.Visible = True
P_P:
   End If
   temp_mayor.MoveNext
   
Loop
i_glosa.SetFocus
If voucher.ItemData(voucher.ListIndex) <> 2 Then
   If i_tipdoc.ListCount > 0 Then i_tipdoc.ListIndex = 0
End If
PSMOV_VOU.rdoParameters(0) = LK_CODCIA
PSMOV_VOU.rdoParameters(1) = LK_FECHA_COP1
PSMOV_VOU.rdoParameters(2) = LK_FECHA_COP2
PSMOV_VOU.rdoParameters(3) = LK_NRO_MES
PSMOV_VOU.rdoParameters(4) = voucher.ItemData(voucher.ListIndex)
VOU_MOV.Requery
If VOU_MOV.EOF Then
   ws_nro_voucher = 0
Else
   ws_nro_voucher = VOU_MOV!MOV_NRO_VOUCHER
End If
ws_nro_voucher = ws_nro_voucher + 1
lbl_nro_voucher.Caption = Format(ws_nro_voucher, "########0.0")
loc_voucher = -1

End Sub

Private Sub i_tipdoc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
i_numser.SetFocus

fin:

End Sub



Private Sub i_moneda_GotFocus()
'If i_moneda.Text = "" Then
'  i_moneda.ListIndex = 0
'Else
'  i_moneda_Click
'End If
End Sub

Private Sub i_moneda_KeyPress(KeyAscii As Integer)
'Dim wtc
'If KeyAscii <> 13 Then Exit Sub
'txtdetalle.SetFocus
End Sub

Private Sub i_tipdoc_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 Then
    PUB_TIPREG = 50
    PUB_CODCIA = "00"
    Load FrmDatArti
    FrmDatArti.Caption = Trim(Left(i_tipdoc.Text, 30)) & " TAB_TIPREG = " & PUB_TIPREG
    FrmDatArti.Show 1
    PUB_TIPREG = 50
    PUB_CODCIA = "00"
    SQ_OPER = 2
    LEER_TAB_LLAVE
    i_tipdoc.Clear
    Do Until tab_mayor.EOF
       i_tipdoc.AddItem Format(tab_mayor!TAB_NUMTAB, "00") & ".-" & tab_mayor!tab_nomlargo
       tab_mayor.MoveNext
    Loop
    If i_tipdoc.ListCount > 0 Then i_tipdoc.ListIndex = 0
    i_tipdoc.SetFocus
    SendKeys "%{UP}"
    DoEvents
End If
End Sub

Private Sub lbl_nro_voucher_DblClick()
Dim wvou
pub_mensaje = "Si desea Modificar en Nro. Correlativo del Voucher seleccione <Si>, de lo contrario siguirá el correlativo <No>"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
no_sabe:
    PSMOV_VOU.rdoParameters(0) = LK_CODCIA
    PSMOV_VOU.rdoParameters(1) = LK_FECHA_COP1
    PSMOV_VOU.rdoParameters(2) = LK_FECHA_COP2
    PSMOV_VOU.rdoParameters(3) = LK_NRO_MES
    PSMOV_VOU.rdoParameters(4) = voucher.ItemData(voucher.ListIndex)
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

Private Sub Option1_Click(Index As Integer)
frmcon.Visible = False
If Index = 1 Then
  frmcon.Visible = True
  If TIPMOV.Visible Then
    TIPMOV.ListIndex = 0
    TIPMOV.SetFocus
  End If
Else
If i_plantilla.Enabled Then i_plantilla.SetFocus
 
End If

End Sub

Private Sub salir_Click()
If LOC_CANCELA = 1 Then
  cmdcorta_Click
  Exit Sub
End If
Unload Me
End Sub


Public Sub LIMPIA_DATOS()
txtdetalle.Text = ""
FORM_CONTA2.grid_fac.Clear
grid_fac.Rows = 3
If voucher.ListIndex >= 0 Then
 If voucher.ItemData(voucher.ListIndex) <> 2 Then
  i_numser.Text = ""
  i_numfac.Text = ""
  i_tipdoc.ListIndex = -1
 End If
End If
lblcodusu.Caption = ""
FLAG_DIF_TC = ""
fila = 1
WPASA = False
End Sub

Public Sub CABE_MOSTRAR()
Grid_cons.Clear

'grid_fac.MergeCells = 4
'grid_fac.MergeCol(4) = True
'grid_fac.MergeCol(5) = True


fila = 0
Grid_cons.Cols = 13

Grid_cons.ColWidth(0) = 0
Grid_cons.ColWidth(1) = 1200
Grid_cons.ColWidth(2) = 1000
Grid_cons.ColWidth(3) = 2000
Grid_cons.ColWidth(4) = 1200
Grid_cons.ColWidth(5) = 1200
Grid_cons.ColWidth(6) = 300
Grid_cons.ColWidth(7) = 1500
Grid_cons.ColWidth(8) = 500
Grid_cons.ColWidth(9) = 900
Grid_cons.ColWidth(10) = 0
Grid_cons.ColWidth(11) = 0
Grid_cons.ColWidth(12) = 2000


Grid_cons.TextMatrix(0, 1) = "Fecha"
Grid_cons.TextMatrix(0, 2) = "Cuenta"
Grid_cons.TextMatrix(0, 3) = "Descrip."
Grid_cons.TextMatrix(0, 4) = "Debe"
Grid_cons.TextMatrix(0, 5) = "Haber"
Grid_cons.TextMatrix(0, 6) = "S/D"
Grid_cons.TextMatrix(0, 7) = "Tipo Doc."
Grid_cons.TextMatrix(0, 8) = "Serie"
Grid_cons.TextMatrix(0, 9) = "N.Docum."
Grid_cons.TextMatrix(0, 12) = "Detalle"


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
   fila = WF - 1
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

Public Sub suma_subtotal()
Exit Sub
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
    SUM_D = SUM_D + Val(grid_fac.TextMatrix(fila, 3))
    SUM_H = SUM_H + Val(grid_fac.TextMatrix(fila, 4))
Next fila
'fila = WF - 1
grid_fac.TextMatrix(WFIN + 1, 3) = Format(SUM_D, "###,##0.00")
grid_fac.TextMatrix(WFIN + 1, 4) = Format(SUM_H, "###,##0.00")
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
      If wsKeyAscii <> 45 And wsKeyAscii <> 8 And wsKeyAscii <> 13 And car <> "." Then
          wsKeyAscii = 0
          Beep
          Exit Sub
        End If
    End If

End Sub

Public Sub CABE_ING()
grid_fac.Cols = 26
grid_fac.Rows = 3
grid_fac.Clear
grid_fac.TextMatrix(grid_fac.Rows - 1, 10) = 1
grid_fac.RowHeight(0) = 285
grid_fac.RowHeight(1) = 285
grid_fac.RowHeight(2) = 285

fila = 0
' falta coddigo sunat doc. relacion.

grid_fac.ColWidth(0) = 400 ' Item
grid_fac.ColWidth(1) = 800 ' Codigo Cuenta
grid_fac.ColWidth(2) = 2200 ' Descripcion de Cuenta
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
grid_fac.ColWidth(20) = 0 ' RUC esposo
grid_fac.ColWidth(21) = 0 ' Descripcion para el Centro de Costo
grid_fac.ColWidth(22) = 0 ' Descripcion para el Codigo de Sunat
grid_fac.ColWidth(23) = 0 ' Descripcion para el OPC
grid_fac.ColWidth(24) = 0 ' tipo de Cambio
grid_fac.ColWidth(25) = 0 ' Flag


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

Private Sub SSALIR_Click()
   Grid_cons.Visible = False
   SSALIR.Visible = False
   ESTADO.Visible = True
   Frame3.Visible = True
   Frame2.Visible = True
   cancelar_Click
End Sub



Public Sub BUSCAR_CTA(WTIPO As Integer)
Dim WCUENTA As TextBox
Dim wgrupo As String
Dim wq_cuenta As String

LK_TABLA = "BUSCAR3"
If WTIPO = 1 Then
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
  Else
  TEXTOVAR2.Text = Trim(frmBuscacta.tcuenta)
  End If
Else
 grid_fac.TextMatrix(grid_fac.Row, 1) = ""
 grid_fac.TextMatrix(grid_fac.Row, 2) = ""
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
  Azul3 TEXTOVAR2, TEXTOVAR2
End If


End Sub



Private Sub TEXTOVAR2_Change()
If TEXTOVAR2.Text = "" Then
   LV_CLI.Visible = False
End If
grid_fac.Text = TEXTOVAR2.Text
If grid_fac.Col = 11 Or grid_fac.Col = 12 Then
  If Val(grid_fac.TextMatrix(grid_fac.Row, 1)) = 0 Then
     TEXTOVAR2.Text = ""
     grid_fac.Text = ""
     Exit Sub
  End If
  If grid_fac.Col = 11 Then
    If Val(grid_fac.TextMatrix(grid_fac.Row, 12)) <> 0 Then
       TEXTOVAR2.Text = "0"
    End If
  End If
  If grid_fac.Col = 12 Then
    If Val(grid_fac.TextMatrix(grid_fac.Row, 11)) <> 0 Then
       TEXTOVAR2.Text = "0"
    End If
  End If
  grid_fac.Text = Format(TEXTOVAR2.Text, "0.00")
  suma_grid
End If

End Sub
Private Sub textovar2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.

If LV_CLI.Visible Then GoTo SALTAX


'If KeyCode = 37 Or KeyCode = 38 Or KeyCode = 39 Or KeyCode = 13 Or KeyCode = 40 Then
'Else
'Exit Sub
'End If
If KeyCode <> 13 Then Exit Sub

TEXTOVAR2.Visible = False
grid_fac.SetFocus

If KeyCode = 38 Then
   If grid_fac.Row = 2 Then Exit Sub
End If
   
If KeyCode = 37 Then
   If grid_fac.Col = 1 Then Exit Sub
End If
If KeyCode = 40 Then
   If grid_fac.Row = grid_fac.Rows - 1 Then Exit Sub
End If
If KeyCode = 13 Then
   If grid_fac.Row = grid_fac.Rows - 1 Then Exit Sub
End If


If KeyCode = 38 Then
   grid_fac.Row = grid_fac.Row - 1
   Exit Sub
ElseIf KeyCode = 40 Then
   grid_fac.Row = grid_fac.Row + 1
   Exit Sub
End If

If KeyCode = 37 Then
   grid_fac.Col = grid_fac.Col - 1
ElseIf KeyCode = 39 Then
   grid_fac.Col = grid_fac.Col + 1
End If

Exit Sub

SALTAX:

If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And TEXTOVAR2.Text = "" Then
  loc_key = 1
  Set LV_CLI.SelectedItem = LV_CLI.ListItems(loc_key)
  LV_CLI.ListItems.Item(loc_key).Selected = True
  LV_CLI.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > LV_CLI.ListItems.Count Then loc_key = LV_CLI.ListItems.Count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > LV_CLI.ListItems.Count Then loc_key = LV_CLI.ListItems.Count
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
  LV_CLI.ListItems.Item(loc_key).Selected = True
  LV_CLI.ListItems.Item(loc_key).EnsureVisible
  TEXTOVAR2.Text = Trim(LV_CLI.ListItems.Item(loc_key).Text) & " "
  DoEvents
  TEXTOVAR2.SelStart = Len(TEXTOVAR2.Text)
  DoEvents
fin:

End Sub
Private Sub textovar2_KeyPress(KeyAscii As Integer)
Dim LISTACC As rdoResultset
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As ListItem    ' Variable FoundItem.
Dim msgerr As String
Dim LenCCgrd As Integer

If KeyAscii = 27 Then
 TEXTOVAR2.Text = ""
 TEXTOVAR2.Visible = False
 grid_fac.SetFocus
 Exit Sub
End If
If grid_fac.Col = 10 And TEXTOVAR2.Text <> "1" And TEXTOVAR2.Text <> "2" And KeyAscii = 13 Then
    TEXTOVAR2.Text = 1
    MsgBox "Valor deber 1 ó 2"
    grid_fac.SetFocus
    Exit Sub
End If
If grid_fac.Col = 1 Then
   If KeyAscii <> 42 Then SOLO_ENTERO KeyAscii
End If
If grid_fac.Col = 2 Then grid_fac.TextMatrix(grid_fac.Row, 19) = ""
If grid_fac.Col = 6 Then grid_fac.TextMatrix(grid_fac.Row, 21) = ""
If grid_fac.Col = 7 Then grid_fac.TextMatrix(grid_fac.Row, 22) = ""
If grid_fac.Col = 5 Then grid_fac.TextMatrix(grid_fac.Row, 23) = ""
    
If grid_fac.Col = 11 Or grid_fac.Col = 12 Then
  SOLO_DECIMAL TEXTOVAR2, KeyAscii
End If

If KeyAscii <> 13 Then
   GoTo fin
End If
If grid_fac.Col = 3 Then GoTo CHEQUEO_CLIENTES
If grid_fac.Col = 6 And Trim(Left(TEXTOVAR2.Text, 1)) = "*" Then ' BUSQUEDA CODIGO DE CENTRO DE COSTOS
  
  pub_cadena = "SELECT * FROM USUARIOS ORDER BY usu_key"
  
  Set usu = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)


  TEXTOVAR2.Text = Mid(TEXTOVAR2.Text, 2, Len(TEXTOVAR2.Text))
  pub_cadena = "SELECT * FROM TABLAS WHERE TAB_TIPREG = 40 AND TAB_CODCIA = '" & LK_CODCIA & "' AND TAB_NOMLARGO LIKE '" & TEXTOVAR2.Text & "%' ORDER BY TAB_NOMLARGO"
  'Debug.Print pub_cadena
  Set LISTACC = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
  LISTACC.Requery
  liscc.Clear
  If LISTACC.EOF Then
     TEXTOVAR2.Visible = True
     TEXTOVAR2.Text = ""
     MsgBox "No Existe Descripción , Verificar ", 48, Pub_Titulo
     TEXTOVAR2.SetFocus
     Exit Sub
  End If
  Do Until LISTACC.EOF
    liscc.AddItem LISTACC!tab_nomlargo & String(80, " ") & LISTACC!TAB_NUMTAB
    LISTACC.MoveNext
  Loop
  liscc.Visible = True
  liscc.SetFocus
  Exit Sub
End If

If grid_fac.Col = 6 Then
    LenCCgrd = Len(Trim(grid_fac.TextMatrix(grid_fac.Row, 6)))
    CC1 = Mid(grid_fac.TextMatrix(grid_fac.Row, 6), 1, PUB_Dgt_CC1)
    CC2 = Mid(grid_fac.TextMatrix(grid_fac.Row, 6), Val(PUB_Dgt_CC1) + 1, Val(PUB_Dgt_CC2))
    CC3 = Mid(grid_fac.TextMatrix(grid_fac.Row, 6), Val(PUB_Dgt_CC1) + Val(PUB_Dgt_CC2) + 1, Val(PUB_Dgt_CC3))
    CC = CC1 & CC2 & CC3
    If Len(CC) = 6 Then
        PSCC_EXIST(0) = CC1
        PSCC_EXIST(1) = CC2
        PSCC_EXIST(2) = CC3
        cc_exist.Requery
        msgerr = ""
        If cc_exist.RowCount <> 3 Then
           Do While Not cc_exist.EOF
              ' msgerr = msgerr & cc_exist("CC_CODIGO") & " - " & cc_exist("CC_DESCRIPCION") & vbCrLf
               cc_exist.MoveNext
           Loop
           MsgBox "Codigo de Centro de Costo No Existe" & vbCrLf & msgerr, 48, Pub_Titulo
           grid_fac.SetFocus
           Exit Sub
        End If
    ElseIf Len(CC) = 0 Then
        GoTo saltacolumna
    Else
        MsgBox "Codigo de Centro de Costo No Existe", 48, Pub_Titulo
        grid_fac.SetFocus
        Exit Sub
    End If
End If
If grid_fac.Col = 7 And Val(TEXTOVAR2.Text) <> 0 Then  ' CODIGO DE TIPO DE DOC. SUNAT
    SQ_OPER = 1
    PUB_CODCIA = "00"
    PUB_TIPREG = 50
    PUB_NUMTAB = Val(TEXTOVAR2.Text)
    LEER_TAB_LLAVE
    If tab_llave.EOF Then
      MsgBox "Codigo de TIPO DE DOC. No Existe", 48, Pub_Titulo
      grid_fac.SetFocus
      grid_fac.TextMatrix(grid_fac.Row, 7) = ""
      Exit Sub
    End If
    grid_fac.TextMatrix(grid_fac.Row, 22) = tab_llave!tab_nomlargo ' Descripcion para el codigo de sunat
End If
'esto bloqueado por mic falta definir que es opc o que dato se puede guardar aca
'************************************************************************************************
'If grid_fac.Col = 5 And Val(TEXTOVAR2.Text) <> 0 Then  ' CODIGO DE TIPO DE DOC. SUNAT
'    SQ_OPER = 1
'    PUB_CODCIA = LK_CODCIA
'    PUB_TIPREG = 240
'    PUB_NUMTAB = Val(TEXTOVAR2.Text)
'    LEER_TAB_LLAVE
'    If tab_llave.EOF Then
'      MsgBox "Codigo de OPC. No Existe", 48, Pub_Titulo
'      grid_fac.SetFocus
'      grid_fac.TextMatrix(grid_fac.Row, 5) = ""
'      Exit Sub
'    End If
'    grid_fac.TextMatrix(grid_fac.Row, 23) = tab_llave!tab_nomlargo ' Descripcion para el codigo de sunat
'End If
'**********************************************************************************************************
'BLOQADO PARA MDIFICACIONES EN PEREDA
'If grid_fac.Col = 9 Then
'    pub_cadena = "SELECT * FROM MoviCont WHERE MOV_CODCIA = '" & LK_CODCIA & "' AND MOV_CODCLIE = " & Val(grid_fac.TextMatrix(grid_fac.Row, 3)) & " AND MOV_SERIE = " & Val(grid_fac.TextMatrix(grid_fac.Row, 8)) & " AND MOV_NUMFAC = " & Val(grid_fac.TextMatrix(grid_fac.Row, 9)) & " AND MOV_CP = '" & grid_fac.TextMatrix(grid_fac.Row, 4) & "' AND MOV_TIPMOV = " & CStr(voucher.ItemData(voucher.ListIndex))
'    Set Doc_exist = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
'    Doc_exist.Requery
'    If Not Doc_exist.EOF Then
'        MsgBox "Documento ya existe"
'        grid_fac.TextMatrix(grid_fac.Row, 9) = ""
'        grid_fac.SetFocus
'    End If
'End If

If grid_fac.Col <> 1 Then
   GoTo saltacolumna
   Exit Sub
End If

cmdIngreso.Visible = True
F2 = 0
On Error GoTo OJO
pu_codclie = Val(TEXTOVAR2.Text)
On Error GoTo 0
If Len(TEXTOVAR2.Text) = 0 Then
   Exit Sub
End If
If pu_codclie <> 0 And IsNumeric(TEXTOVAR2.Text) = True Then
   On Error GoTo OJO
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   PUB_CUENTA = TEXTOVAR2.Text
   PUB_CODCIA = LK_CODCIA
   LEER_COM_LLAVE
   On Error GoTo 0
   If com_llave.EOF Then
    MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
    TEXTOVAR2.Visible = True
    Azul TEXTOVAR2, TEXTOVAR2
    GoTo fin
   Else
'bloqueado por mic para para q permita pasar los diferentes niveles
     If Val(com_llave!com_flag_afectacion) <> 1 Then
'mic If Val(com_llave!com_nivel) <> cop_llave!cop_nivel_max Then
        MsgBox "No es Cuenta Analitica ...", 48, Pub_Titulo
        TEXTOVAR2.Visible = True
        Azul TEXTOVAR2, TEXTOVAR2
        'TEXTOVAR2.SetFocus
        GoTo fin
     Else
        grid_fac.TextMatrix(grid_fac.Row, 2) = Trim(com_llave!com_DESCRIPCION)
        GoTo saltacolumna
        'If Trim(grid_fac.TextMatrix(grid_fac.Row, 11)) = "D" Then
        '    grid_fac.COL = 3
        'Else
        '    grid_fac.COL = 4
        'End If
'       grid_fac.Col = 3
        TEXTOVAR2.Visible = False
        Exit Sub
        If Not TEXTOVAR.Visible Then grid_fac.SetFocus
     End If
   End If
Else
On Error GoTo SIGUE

If Left(TEXTOVAR2.Text, 1) = "*" Then
  TEXTOVAR.Text = TEXTOVAR2.Text
  BUSCAR_CTA 0
  Exit Sub
End If

   If loc_key <> 0 Then valor = UCase(LV_CLI.ListItems.Item(loc_key).Text)
   If Trim(UCase(TEXTOVAR2.Text)) = Left(valor, Len(Trim(TEXTOVAR2.Text))) Then
   Else
      Exit Sub
   End If
   If loc_key = 0 Then Exit Sub
   TEXTOVAR2.Text = Trim(LV_CLI.ListItems.Item(loc_key).SubItems(1))
   PUB_CUENTA = Val(TEXTOVAR2.Text)
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   PUB_CODCIA = LK_CODCIA
   LEER_COM_LLAVE
End If



LV_CLI.Visible = False
grid_fac.TextMatrix(grid_fac.Row, 2) = Trim(com_llave!com_DESCRIPCION)
grid_fac.Col = 3
TEXTOVAR2.Visible = False
If Not TEXTOVAR.Visible Then grid_fac.SetFocus

Exit Sub
SIGUE:
If Err.Number = 35600 Then
  Exit Sub
End If
fin:
OJO:

Exit Sub

'*********
saltacolumna:
If grid_fac.Col = 1 Then
   grid_fac.Col = 3
ElseIf grid_fac.Col = 3 Then
   grid_fac.Col = 5
ElseIf grid_fac.Col = 5 Then
   grid_fac.Col = 6
ElseIf grid_fac.Col = 6 Then
   grid_fac.Col = 7
ElseIf grid_fac.Col = 7 Then
   grid_fac.Col = 8
ElseIf grid_fac.Col = 8 Then
   grid_fac.Col = 9
ElseIf grid_fac.Col = 9 Then
   grid_fac.Col = 10
ElseIf grid_fac.Col = 10 Then
   grid_fac.Col = 11
ElseIf grid_fac.Col = 11 Then
   grid_fac.Col = 12
ElseIf grid_fac.Col = 12 Then
  If grid_fac.Row = grid_fac.Rows - 1 Then
    grid_fac.Rows = grid_fac.Rows + 1
    grid_fac.TextMatrix(grid_fac.Rows - 1, 10) = 1
    grid_fac.Row = grid_fac.Rows - 1
    grid_fac.RowHeight(grid_fac.Rows - 1) = 300
  Else
    grid_fac.Row = grid_fac.Row + 1
  End If
   grid_fac.Col = 1
End If
   grid_fac.SetFocus
   
Exit Sub
' PARA CHEQUE DE CLIENTES PROVEEDRES Y PERSONAL
CHEQUEO_CLIENTES:
On Error GoTo OJO
pu_codclie = Val(TEXTOVAR2.Text)
On Error GoTo 0
If Len(TEXTOVAR2.Text) = 0 Then
   'Exit Sub
   GoTo saltacolumna
End If
If pu_codclie <> 0 And IsNumeric(TEXTOVAR2.Text) = True Then
   On Error GoTo OJO
   
   SQ_OPER = 1
   pu_cp = Trim(grid_fac.TextMatrix(grid_fac.Row, 4))
   If pu_cp = "E" Then
      PSPER_BUSCA(0) = Val(TEXTOVAR2.Text)
      per_busca.Requery
      If per_busca.EOF Then
          Azul TEXTOVAR2, TEXTOVAR2
          MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
          TEXTOVAR2.Visible = True
          TEXTOVAR2.SetFocus
          GoTo fin
       Else
          grid_fac.SetFocus
          Exit Sub
       End If
   Else
      pu_codcia = LK_CODCIA
      PUB_CODCLIE = Val(TEXTOVAR2.Text)
      PUB_RUC = PUB_CODCLIE
      On Error GoTo 0
      If Len(TEXTOVAR2.Text) = 11 Then
        SQ_OPER = 4
        LEER_CLI_LLAVE
        SQ_OPER = 1
        If Not cli_ruc.EOF Then
            PUB_CODCLIE = cli_ruc!CLI_CODCLIE
            pu_codclie = PUB_CODCLIE
        Else
            PUB_CODCLIE = ""
            pu_codclie = ""
        End If
      End If
        LEER_CLI_LLAVE
        If cli_llave.EOF Then
            MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
            TEXTOVAR2.Visible = True
            Azul TEXTOVAR2, TEXTOVAR2
            GoTo fin
        Else
            grid_fac.SetFocus
            grid_fac.TextMatrix(grid_fac.Row, 19) = cli_llave!cli_nombre ' Descripcion para el Codigo
            grid_fac.TextMatrix(grid_fac.Row, 3) = pu_codclie
            GoTo saltacolumna
            Exit Sub
        End If
   End If
Else
   On Error GoTo SIGUE
   If loc_key <> 0 Then valor = UCase(LV_CLI.ListItems.Item(loc_key).Text)
   If Trim(UCase(TEXTOVAR2.Text)) = Left(valor, Len(Trim(TEXTOVAR2.Text))) Then
   Else
      Exit Sub
   End If
   If loc_key = 0 Then Exit Sub
   TEXTOVAR2.Text = Trim(LV_CLI.ListItems.Item(loc_key).SubItems(1))
   pu_cp = grid_fac.TextMatrix(grid_fac.Row, 4)
   pu_codcia = LK_CODCIA
   pu_codclie = TEXTOVAR2.Text
   LEER_CLI_LLAVE
   grid_fac.TextMatrix(grid_fac.Row, 19) = Trim(LV_CLI.ListItems.Item(loc_key).Text) ' .SubItems(0))
   grid_fac.TextMatrix(grid_fac.Row, 20) = Trim(cli_llave!cli_ruc_esposo)
End If

LV_CLI.Visible = False
TEXTOVAR2.Visible = False
grid_fac.SetFocus
Exit Sub


End Sub

Private Sub textovar2_KeyUp(KeyCode As Integer, Shift As Integer)

Dim itmFound As ListItem    ' Variable FoundItem.
Dim var, VAR2
Dim LenCCgrd  As Integer

If KeyCode = 112 And grid_fac.Col = 6 Then
    fra_cc.Visible = True
    LenCCgrd = Len(Trim(grid_fac.TextMatrix(grid_fac.Row, 6)))
    CC1 = Mid(grid_fac.TextMatrix(grid_fac.Row, 6), 1, Val(PUB_Dgt_CC1))
    CC2 = Mid(grid_fac.TextMatrix(grid_fac.Row, 6), Val(PUB_Dgt_CC1) + 1, Val(PUB_Dgt_CC2))
    CC3 = Mid(grid_fac.TextMatrix(grid_fac.Row, 6), Val(PUB_Dgt_CC1) + Val(PUB_Dgt_CC2) + 1, Val(PUB_Dgt_CC3))
    CC = CStr(CC1) + CStr(CC2) + CStr(CC3)
    If CC <> "" Then
        If lstcc(0).Enabled Then FindCC 0, Val(CC1)
        If lstcc(1).Enabled Then FindCC 1, Val(CC2)
        If lstcc(2).Enabled Then FindCC 2, Val(CC3)
    End If
    fra_cc.Caption = "Centro de Costo: " + CC
    lstcc(0).SetFocus
    Exit Sub
End If


If grid_fac.Col = 3 Then
If Len(TEXTOVAR2.Text) = 0 Or IsNumeric(TEXTOVAR2.Text) Then
   LV_CLI.Visible = False
   Exit Sub
End If
'If LV_CLI.Visible = False And KeyCode <> 13 Or Len(i_codcli.Text) = 1 Then
If LV_CLI.Visible = False And KeyCode <> 13 Or Len(Trim(TEXTOVAR2.Text)) = 1 Then
   'grid_fac.TextMatrix(grid_fac.Row, 4) = "C"
    XX_CUENTA = Trim(grid_fac.TextMatrix(grid_fac.Row, 4))
    loc_key = 0
    var = Asc(TEXTOVAR2.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    VAR2 = Val(XX_CUENTA) + 1
    
    If Trim(XX_CUENTA) <> "" Then
    
       PUB_CP = XX_CUENTA
       If PUB_CP = "E" Then
         numarchi = 2
         archi = "SELECT PER_INDEX, PER_CODCIA,  [Nombres y Apellidos]  FROM LISTAPERSONAL  WHERE PER_CODCIA IN ('01','02','03','04') AND [Nombres y Apellidos] BETWEEN '" & TEXTOVAR2.Text & "' AND  '" & var & "' ORDER BY [Nombres y Apellidos]"
       Else
         numarchi = 1
         archi = "SELECT CLI_CODCLIE , CLI_CODCIA, CLI_CP, CLI_NOMBRE, CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM, CLI_RUC_ESPOSO, TAB_NOMLARGO  FROM CLIENTES,TABLAS WHERE (TAB_CODCIA = '00') AND (TAB_TIPREG = 35) AND (TAB_NUMTAB = CLI_ZONA_NEW) AND CLI_CP = '" & PUB_CP & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & TEXTOVAR2.Text & "' AND  '" & var & "' ORDER BY CLI_NOMBRE"
       End If
       
    End If
       
    PROC_LISVIEW LV_CLI
    LV_CLI.Top = 700
    LV_CLI.Height = 2500
    LV_CLI.Width = 8000
    LV_CLI.Left = 3500
    LV_CLI.ZOrder 0

    loc_key = 0
    If LV_CLI.Visible Then
     loc_key = 1
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If

If LV_CLI.Visible Then
  Set itmFound = LV_CLI.FindItem(LTrim(TEXTOVAR2.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = Val(itmFound.Tag)
   If loc_key + 8 > LV_CLI.ListItems.Count Then
      LV_CLI.ListItems.Item(LV_CLI.ListItems.Count).EnsureVisible
   Else
     LV_CLI.ListItems.Item(loc_key + 8).EnsureVisible
   End If
  End If
  Exit Sub
End If



End If

If grid_fac.Col <> 1 Then Exit Sub
If Len(TEXTOVAR2.Text) = 0 Or IsNumeric(TEXTOVAR2.Text) Then
   LV_CLI.Visible = False
   Exit Sub
End If
If LV_CLI.Visible = False Or Len(Trim(TEXTOVAR2.Text)) = 1 Then
   loc_key = 0
    var = Asc(TEXTOVAR2.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    VAR2 = Val(XX_CUENTA) + 1
    
    numarchi = 5
    archi = "SELECT COM_CUENTA , COM_CODCIA, COM_DESCRIPCION, COM_NIVEL   FROM COMAEST WHERE COM_CODCIA = '" & LK_CODCIA & "' AND COM_DESCRIPCION BETWEEN '" & TEXTOVAR2.Text & "' AND  '" & var & "'  AND COM_CUENTA BETWEEN '" & XX_CUENTA & "' AND  '" & VAR2 & "' AND COM_NIVEL =" & cop_llave!cop_nivel_max & "  ORDER BY COM_DESCRIPCION"
       
    PROC_LISVIEW LV_CLI
    loc_key = 0
    If LV_CLI.Visible Then
     loc_key = 1
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
If LV_CLI.Visible Then
  Set itmFound = LV_CLI.FindItem(LTrim(TEXTOVAR2.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = Val(itmFound.Tag)
   If loc_key + 8 > LV_CLI.ListItems.Count Then
      LV_CLI.ListItems.Item(LV_CLI.ListItems.Count).EnsureVisible
   Else
     LV_CLI.ListItems.Item(loc_key + 8).EnsureVisible
   End If
  End If
  Exit Sub
End If

End Sub

Private Sub TEXTOVAR2_LostFocus()
TEXTOVAR2.Visible = False
End Sub

Private Sub TIPMOV_Click()
TIPMOV_KeyPress 13
End Sub

Private Sub TIPMOV_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub

PSMOV_VOU.rdoParameters(0) = LK_CODCIA
PSMOV_VOU.rdoParameters(1) = LK_FECHA_COP1
PSMOV_VOU.rdoParameters(2) = LK_FECHA_COP2
PSMOV_VOU.rdoParameters(3) = LK_NRO_MES
PSMOV_VOU.rdoParameters(4) = Val(TIPMOV.Text)
VOU_MOV.Requery
If VOU_MOV.EOF Then
   ws_nro_voucher = 0
Else
   ws_nro_voucher = VOU_MOV!MOV_NRO_VOUCHER
End If
 tvou.Text = Format(ws_nro_voucher, "########")
 tvou.SetFocus
 tvou_KeyPress 13
End Sub

Private Sub tvou_KeyPress(KeyAscii As Integer)
On Error GoTo SALE
If KeyAscii <> 13 Then Exit Sub
PSMOV_LLAVE.rdoParameters(0) = LK_CODCIA
PSMOV_LLAVE.rdoParameters(1) = Val(tvou.Text)
PSMOV_LLAVE.rdoParameters(2) = Val(TIPMOV.Text)
PSMOV_LLAVE.rdoParameters(3) = LK_FECHA_COP1
PSMOV_LLAVE.rdoParameters(4) = LK_FECHA_COP2
mov_llave.Requery
men.Caption = ""
If mov_llave.EOF Then
'   MsgBox "Voucher no Existe...", 48, Pub_Titulo
   LIMPIA_DATOS
   lbl_nro_voucher.Caption = Trim(tvou.Text)
   men.Caption = "Documento en Blanco"
   i_plantilla.ListIndex = -1
   voucher.ListIndex = -1
   i_glosa.Text = ""
   Exit Sub
Else
   DoEvents
   pb.Visible = True
   DoEvents
   pb.Min = 0
   pb.Value = 0
   pb.Max = mov_llave.RowCount
   
   CABE_ING
   'Grid_cons.Visible = False
   SSALIR.Visible = False
   ESTADO.Visible = True
   Frame3.Visible = True
   GoSub busca_fecha
   GoSub busca_voucher
   GoSub busca_PLAN
   'Voucher_LostFocus
   If mov_llave!MOV_MONEDA = "S" Then i_moneda.ListIndex = 0
   'If mov_llave!MOV_MONEDA = "D" Then i_moneda.ListIndex = 1
   GoSub busca_tipdoc
   i_numser.Text = mov_llave!MOV_serie
   i_numfac.Text = mov_llave!MOV_numfac
   i_glosa.Text = Trim(mov_llave!MOV_GLOSA)
   txtdetalle.Text = Trim(mov_llave!MOV_DETALLE)
   lbl_nro_voucher.Caption = Format(mov_llave!MOV_NRO_VOUCHER, "#######0.0")
   loc_voucher = mov_llave!MOV_NRO_VOUCHER
   lblcodusu.Caption = mov_llave!MOV_CODUSU
   If mov_llave!MOV_FLAG_DES = "A" Then
     lbldes.Caption = "Asiento de Destino."
     destino.Caption = "Re&gresar"
   Else
     lbldes.Caption = ""
     destino.Caption = "&Ver Destino"
   End If
   fila = 1
   Do Until mov_llave.EOF
    pb.Value = pb.Value + 1
    fila = fila + 1
    grid_fac.Rows = fila + 1
    grid_fac.RowHeight(grid_fac.Rows - 1) = 285
    grid_fac.TextMatrix(grid_fac.Rows - 1, 0) = Format(fila, "00")
    grid_fac.TextMatrix(grid_fac.Rows - 1, 1) = Trim(mov_llave!MOV_CODCTA)
    SQ_OPER = 1
    PUB_CUENTA = mov_llave!MOV_CODCTA
    PUB_CODCIA = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
      grid_fac.TextMatrix(grid_fac.Rows - 1, 2) = "Cuenta Invalidad .."
    Else
      grid_fac.TextMatrix(grid_fac.Rows - 1, 2) = Trim(com_llave!com_DESCRIPCION)
    End If
    If mov_llave!MOV_DH = "D" Then
      grid_fac.TextMatrix(grid_fac.Rows - 1, 11) = mov_llave!MOV_IMPORTE
    Else
      grid_fac.TextMatrix(grid_fac.Rows - 1, 12) = mov_llave!MOV_IMPORTE
    End If
    If Val(mov_llave!MOV_SUNAT) <> 0 Then
      grid_fac.TextMatrix(grid_fac.Rows - 1, 7) = Format(mov_llave!MOV_SUNAT, "00")
    Else
      grid_fac.TextMatrix(grid_fac.Rows - 1, 7) = Format(mov_llave!MOV_SUNAT, "#")
    End If
    If Val(Nulo_Valor0(mov_llave!MOV_OPC)) <> 0 Then
      grid_fac.TextMatrix(grid_fac.Rows - 1, 5) = Format(mov_llave!MOV_OPC, "00")
    Else
      grid_fac.TextMatrix(grid_fac.Rows - 1, 5) = Format(mov_llave!MOV_OPC, "#")
    End If
    If Val(mov_llave!MOV_serie) <> 0 Then
      grid_fac.TextMatrix(grid_fac.Rows - 1, 8) = Format(mov_llave!MOV_serie, "000")
    Else
      grid_fac.TextMatrix(grid_fac.Rows - 1, 8) = Format(mov_llave!MOV_serie, "#")
    End If
    If Val(mov_llave!MOV_numfac) <> 0 Then
      grid_fac.TextMatrix(grid_fac.Rows - 1, 9) = Format(mov_llave!MOV_numfac, "0000000")
    Else
      grid_fac.TextMatrix(grid_fac.Rows - 1, 9) = Format(mov_llave!MOV_numfac, "#")
    End If
    grid_fac.TextMatrix(grid_fac.Rows - 1, 3) = Format(mov_llave!MOV_codclie, "#")
    grid_fac.TextMatrix(grid_fac.Rows - 1, 4) = mov_llave!MOV_CP
    If grid_fac.TextMatrix(grid_fac.Rows - 1, 4) = "C" Then
        grid_fac.TextMatrix(1, 3) = "Cliente"
    ElseIf grid_fac.TextMatrix(grid_fac.Rows - 1, 4) = "P" Then
        grid_fac.TextMatrix(1, 3) = "Proveedor"
    ElseIf grid_fac.TextMatrix(grid_fac.Rows - 1, 4) = "E" Then
        grid_fac.TextMatrix(1, 3) = "Personal"
    End If
    
    
    

    'kpMOV_FBG = " "
    'kpMOV_MARCA = " "
    'kpMOV_FBG_C = " "
    'kpMOV_SERIE_C = 0
    'kpMOV_NUMFAC_C = 0
    'kpMOV_FLAG_TC = " "
    'kpMOV_TIPO_CAMBIO = 0 ' grid_fac.TextMatrix(fila_cont, 1)
    'kpMOV_FLAG_DES = " " ' grid_fac.TextMatrix(fila_cont, 1)
    grid_fac.TextMatrix(grid_fac.Rows - 1, 6) = mov_llave!MOV_CC
 '   grid_fac.TextMatrix(grid_fac.Rows - 1, 5) = Format(mov_llave!MOV_OPC, "#")
    grid_fac.TextMatrix(grid_fac.Rows - 1, 10) = Format(mov_llave!MOV_EXONERADO, "0")
    
        SQ_OPER = 1
        PUB_TIPREG = 50
        PUB_NUMTAB = mov_llave("MOV_SUNAT")
        PUB_CODCIA = "00"
        LEER_TAB_LLAVE
        If Not tab_llave.EOF Then
            grid_fac.TextMatrix(grid_fac.Rows - 1, 22) = Nulo_Valors(tab_llave("Tab_NomLargo"))
        End If
'      grid_fac.RowHeight(grid_fac.Rows - 1) = 285
'      grid_fac.TextMatrix(FILA, 1) = Trim(mov_llave!MOV_CODCTA)
'      grid_fac.TextMatrix(FILA, 11) = mov_llave!MOV_DH
'      SQ_OPER = 1
'      PUB_CUENTA = mov_llave!MOV_CODCTA
'      pu_codcia = LK_CODCIA
'      PUB_CODCIA = LK_CODCIA
'      LEER_COM_LLAVE
'      If Not com_llave.EOF Then grid_fac.TextMatrix(FILA, 2) = Trim(com_llave!com_descripcion)
'      If Val(com_llave!COM_NIVEL) <> cop_llave!cop_nivel_max Then grid_fac.TextMatrix(FILA, 2) = ""
'      If mov_llave!MOV_DH = "D" Then
'         grid_fac.TextMatrix(FILA, 3) = mov_llave!MOV_IMPORTE
'      Else
'         grid_fac.TextMatrix(FILA, 4) = mov_llave!MOV_IMPORTE
'      End If
      If Nulo_Valor0(mov_llave!MOV_codclie) <> 0 Then
        SQ_OPER = 1
        pu_cp = Nulo_Valors(mov_llave!MOV_CP)
        pu_codcia = LK_CODCIA
        pu_codclie = Val(mov_llave!MOV_codclie)
        LEER_CLI_LLAVE
        If Not cli_llave.EOF Then
           grid_fac.TextMatrix(fila, 19) = Nulo_Valor0(cli_llave!cli_nombre)
'           grid_fac.TextMatrix(fila, 6) = cli_llave!MOV_CODCLIE
'           grid_fac.TextMatrix(fila, 7) = mov_llave!MOV_CP
'           grid_fac.TextMatrix(fila, 9) = mov_llave!MOV_CP
        End If
      End If
'
'      If Val(mov_llave!MOV_SERIE_C) = 0 And Val(mov_llave!MOV_NUMFAC_C) = 0 Then
'      Else
'       grid_fac.TextMatrix(FILA, 8) = Format(mov_llave!MOV_FBG_C, "00") & "-" & Format(mov_llave!MOV_SERIE_C, "000") & "-" & mov_llave!MOV_NUMFAC_C
'      End If
       If Nulo_Valor0(mov_llave!MOV_TIPO_CAMBIO) <> 0 Then grid_fac.TextMatrix(fila, 24) = Format(mov_llave!MOV_TIPO_CAMBIO, "0.0000")
       grid_fac.TextMatrix(fila, 25) = Nulo_Valors(mov_llave!MOV_FLAG_TC)

       mov_llave.MoveNext
   Loop
End If


  
pb.Visible = False
DoEvents
suma_grid
suma_subtotal
grid_fac.Row = 2
grid_fac.Col = 2
cmdIngreso.Visible = True
grid_fac.SetFocus
Exit Sub
   
   
   
   

busca_fecha:
i_fecha2.Text = Format(mov_llave!MOV_fecha_EMI, "dd/mm/yyyy")
Return

busca_tipdoc:
fila = 0
Do Until i_tipdoc.ListCount - 1 = fila
   i_tipdoc.ListIndex = fila
   If Val(mov_llave!MOV_SUNAT) = Val(i_tipdoc.Text) Then
     Return
   End If
   i_tipdoc.ListIndex = fila
   fila = fila + 1
Loop
If i_tipdoc.ListCount > 0 Then i_tipdoc.ListIndex = 0
Return

busca_voucher:
fila = 0
Do Until voucher.ItemData(voucher.ListIndex) = Val(mov_llave!MOV_TIPMOV) Or fila > 100
   voucher.ListIndex = fila
   fila = fila + 1
Loop
If fila > 100 Then MsgBox "Error de Voucher..."
Return

busca_PLAN:
fila = 0
'On Error GoTo S001
For fila = 0 To i_plantilla.ListCount - 1
 i_plantilla.ListIndex = fila
 If Val(Right(i_plantilla.Text, 5)) = Val(Nulo_Valor0(mov_llave!MOV_PLANTILLA)) Then
    Exit For
 End If
Next fila
'If fila > 100 Then MsgBox "Error de Voucher..."
Return




Exit Sub
SALE:
Resume Next
Exit Sub
S001:
Resume Next

End Sub



Private Sub txtdetalle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If grid_fac.Rows > 2 And grid_fac.Cols > 2 Then
       grid_fac.Row = 2
       If Trim(grid_fac.TextMatrix(grid_fac.Row, 11)) = "D" Then
           grid_fac.Col = 10
       Else
           grid_fac.Col = 11
       End If
       grid_fac_EnterCell
    '  grid_fac.SetFocus
    End If
End If

End Sub

Private Sub Voucher_Click()
If voucher.ListIndex = -1 Then Exit Sub
PSTEMP_MAYOR.rdoParameters(0) = LK_CODCIA
PSTEMP_MAYOR.rdoParameters(1) = voucher.ItemData(voucher.ListIndex)
PSTEMP_MAYOR.rdoParameters(2) = 0
temp_mayor.Requery
If temp_mayor.EOF Then
'   MsgBox "No hay Plantillas"
   Exit Sub
End If
   i_plantilla.Clear
   Do Until temp_mayor.EOF
      If temp_mayor!PLT_SECUENCIA = 0 Then
         i_plantilla.AddItem temp_mayor!PLT_NOMBRE & String(80, "  ") & temp_mayor!PLT_NUMERO
      End If
      temp_mayor.MoveNext
   Loop
If i_plantilla.ListCount > 0 Then i_plantilla.ListIndex = 0

'i_plantilla.Clear
'CABE_ING
End Sub

Private Sub Voucher_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
i_plantilla.SetFocus

End Sub

Public Sub CABE_DOCU()
grid_docu.Cols = 11
grid_docu.Rows = 2
grid_docu.Clear
grid_docu.RowHeight(0) = 285
grid_docu.ColWidth(0) = 900
grid_docu.ColWidth(1) = 600
grid_docu.ColWidth(2) = 1000
grid_docu.ColWidth(3) = 1500
grid_docu.ColWidth(4) = 1100
grid_docu.ColWidth(5) = 1100
grid_docu.ColWidth(6) = 1100
grid_docu.ColWidth(7) = 0
grid_docu.ColWidth(8) = 0
grid_docu.ColWidth(9) = 0
grid_docu.ColWidth(10) = 0 ' moneda
grid_docu.TextMatrix(0, 0) = "Fecha"
grid_docu.TextMatrix(0, 1) = "Mes"
grid_docu.TextMatrix(0, 2) = "Nro.Vou."
grid_docu.TextMatrix(0, 3) = "Documento"
grid_docu.TextMatrix(0, 4) = "Debe"
grid_docu.TextMatrix(0, 5) = "Haber"
grid_docu.TextMatrix(0, 6) = "Saldo"
grid_docu.TextMatrix(0, 7) = ""
grid_docu.TextMatrix(0, 8) = ""
grid_docu.TextMatrix(0, 9) = ""
grid_docu.TextMatrix(0, 10) = ""

End Sub

Private Sub voucher_LostFocus()
Dim IdVoucher As Integer
If voucher.ListIndex >= 0 Then
    IdVoucher = voucher.ItemData(voucher.ListIndex)
End If
    
PSMOV_VOU.rdoParameters(0) = LK_CODCIA
PSMOV_VOU.rdoParameters(1) = LK_FECHA_COP1
PSMOV_VOU.rdoParameters(2) = LK_FECHA_COP2
PSMOV_VOU.rdoParameters(3) = LK_NRO_MES
PSMOV_VOU.rdoParameters(4) = IdVoucher
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
'================================
'================================
Private Sub tvou_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        fra_cc.Visible = True
    End If
End Sub


Private Sub lstcc_KeyPress(Index As Integer, KeyAscii As Integer)
    lstcc_Change Index
    If Index < 2 Then
        If lstcc(Index + 1).Enabled Then lstcc(Index + 1).SetFocus
    End If
    If KeyAscii = 13 And ((Index = 2 And PUB_Flag_CC3 = 1) Or (Index = 1 And Not lstcc(2).Enabled And PUB_Flag_CC2 = 1) Or (Index = 0 And Not lstcc(1).Enabled And PUB_Flag_CC1 = 1)) Then
        grid_fac.TextMatrix(grid_fac.Row, 6) = CC
        GoTo SALIR
    End If
    If KeyAscii = 27 Then
        GoTo SALIR
    End If
Exit Sub
SALIR:
    fra_cc.Visible = False
    grid_fac.Col = 7
    grid_fac.SetFocus
End Sub
Private Sub lstcc_Change(Index As Integer)
On Error GoTo Handler
Select Case Index
        Case 0
                CC1 = ""
                CC1 = lstcc(Index).ItemData(lstcc(Index).ListIndex)
                CC1 = Format(CC1, "000000")
                CC1 = Mid(CC1, 7 - PUB_Dgt_CC1, PUB_Dgt_CC1)
        Case 1
                CC2 = ""
                CC2 = lstcc(Index).ItemData(lstcc(Index).ListIndex)
                CC2 = Format(CC2, "000000")
                CC2 = Mid(CC2, 7 - PUB_Dgt_CC2, PUB_Dgt_CC2)
        Case 2
                CC3 = ""
                CC3 = lstcc(Index).ItemData(lstcc(Index).ListIndex)
                CC3 = Format(CC3, "000000")
                CC3 = Mid(CC3, 7 - PUB_Dgt_CC3, PUB_Dgt_CC3)
    End Select
    CC = CC1 & CC2 & CC3
    fra_cc.Caption = "Centro de Costo: " + CC
    If Len(CC) <> (Val(PUB_Dgt_CC1) + Val(PUB_Dgt_CC2) + Val(PUB_Dgt_CC3)) Then GoTo Handler
    
Exit Sub
Handler:
    CC = ""
End Sub
Private Sub FindCC(ByVal Index As Integer, ByVal valor As Integer)
Dim i As Integer
    For i = 0 To lstcc(Index).ListCount - 1
        If lstcc(Index).ItemData(i) = valor Then
            lstcc(Index).ListIndex = i
            Exit Sub
        End If
    Next i
    lstcc(Index).ListIndex = -1
End Sub
