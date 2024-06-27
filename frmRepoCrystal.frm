VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form RCRYSTAL 
   Caption         =   "Listado en Crystal Report"
   ClientHeight    =   7740
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11565
   ControlBox      =   0   'False
   Icon            =   "frmRepoCrystal.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   11565
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtporcentaje 
      Height          =   285
      Left            =   1440
      TabIndex        =   112
      Top             =   5655
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.ComboBox cboUsers 
      Height          =   315
      Left            =   1335
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   110
      Top             =   6015
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.ComboBox cboDias 
      Height          =   315
      Left            =   1350
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   108
      Top             =   5640
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Frame fraflag 
      Height          =   375
      Left            =   7320
      TabIndex        =   58
      Top             =   4320
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CheckBox cheflag 
         Caption         =   "Check1"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   120
         Width           =   2655
      End
   End
   Begin VB.Frame FRAFLAG2 
      Height          =   375
      Left            =   7320
      TabIndex        =   106
      Top             =   4680
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CheckBox CHEFLAG2 
         Caption         =   "Check1"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   107
         Top             =   120
         Width           =   2655
      End
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   495
      Left            =   8880
      TabIndex        =   22
      Top             =   7080
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   128
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ListView ListView2 
      Height          =   375
      Left            =   8040
      TabIndex        =   26
      Top             =   7080
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   128
      BackColor       =   14737632
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ListView ListView3 
      Height          =   375
      Left            =   7200
      TabIndex        =   34
      Top             =   7080
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
      ForeColor       =   128
      BackColor       =   14737632
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame fratipo 
      Caption         =   "Rubro"
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   7320
      TabIndex        =   97
      Top             =   5880
      Visible         =   0   'False
      Width           =   2895
      Begin VB.ListBox lsttipo 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   510
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   98
         Top             =   240
         Width           =   2655
      End
   End
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   210
      Left            =   0
      TabIndex        =   96
      Top             =   6840
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame frapasa 
      BackColor       =   &H008B4914&
      Caption         =   "Contabilidad"
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   10440
      TabIndex        =   88
      Top             =   0
      Visible         =   0   'False
      Width           =   1300
      Begin VB.CheckBox chepasa 
         BackColor       =   &H008B4914&
         Caption         =   "<--Marcar"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   89
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblpasa2 
         BackColor       =   &H008B4914&
         Caption         =   "al"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   93
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label cop_fecha2 
         BackColor       =   &H008B4914&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label cop_fecha1 
         BackColor       =   &H008B4914&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblpasa 
         BackColor       =   &H008B4914&
         Caption         =   "Pasar la Información al Periodo Contable"
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
         Height          =   855
         Left            =   120
         TabIndex        =   90
         Top             =   240
         Width           =   1050
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Consolidar Compañias :"
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
      Height          =   1095
      Left            =   5760
      TabIndex        =   71
      Top             =   0
      Width           =   4455
      Begin VB.ListBox LISCIA 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   735
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   72
         Top             =   240
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label lblciaact 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   95
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Frame fra1 
      Caption         =   "Descripción del Reporte :"
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   45
      TabIndex        =   3
      Top             =   0
      Width           =   5655
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Archivo: "
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   76
         Top             =   840
         Width           =   630
      End
      Begin VB.Label LBLRUTA 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   840
         TabIndex        =   75
         Top             =   870
         Width           =   4710
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Formula:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   74
         Top             =   600
         Width           =   600
      End
      Begin VB.Label lblformulas 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   840
         TabIndex        =   73
         Top             =   630
         Width           =   4695
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblreporte 
         Alignment       =   2  'Center
         Caption         =   "Listado de Articulos"
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
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame frag 
      Height          =   1335
      Left            =   5520
      TabIndex        =   60
      Top             =   1920
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox cta3 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   720
         TabIndex        =   63
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox cta2 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   720
         TabIndex        =   62
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox cta1 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   720
         TabIndex        =   61
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lfil 
         Caption         =   "Cta. 1:"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   69
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lfil 
         Caption         =   "Cta. 2:"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   68
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lfil 
         Caption         =   "Cta. 3:"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   67
         Top             =   840
         Width           =   615
      End
      Begin VB.Label ncta1 
         Caption         =   "."
         Height          =   255
         Left            =   1920
         TabIndex        =   66
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label ncta2 
         Caption         =   "."
         Height          =   255
         Left            =   1920
         TabIndex        =   65
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label ncta3 
         Caption         =   "."
         Height          =   255
         Left            =   1920
         TabIndex        =   64
         Top             =   840
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Banco :"
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   5760
      TabIndex        =   31
      Top             =   1080
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox txt_key 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   120
         MaxLength       =   8
         TabIndex        =   32
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblbanco 
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
         Left            =   1680
         TabIndex        =   33
         Top             =   240
         Width           =   2565
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame frafechas 
      Caption         =   "Fechas :"
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   45
      TabIndex        =   15
      Top             =   1065
      Visible         =   0   'False
      Width           =   5655
      Begin MSMask.MaskEdBox txtCampo2 
         Height          =   285
         Left            =   3840
         TabIndex        =   16
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   14737632
         ForeColor       =   128
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
      Begin MSMask.MaskEdBox txtCampo1 
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   14737632
         ForeColor       =   128
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
      Begin VB.Label lblcampo1 
         Caption         =   "Campo1"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   945
      End
      Begin VB.Label lblcampo2 
         Caption         =   "Campo1"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2400
         TabIndex        =   18
         Top             =   240
         Width           =   1305
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nº de formulas :"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   8280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame fracal 
      Caption         =   "Calidad del Producto :"
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   0
      TabIndex        =   38
      Top             =   5880
      Visible         =   0   'False
      Width           =   2535
      Begin VB.ListBox listacal 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   510
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   39
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame rango 
      Caption         =   "Rangos  del    :"
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   120
      TabIndex        =   40
      Top             =   5880
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox op2 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1440
         TabIndex        =   42
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox op1 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   240
         TabIndex        =   41
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraven 
      Caption         =   "Lista de Vendedores"
      ForeColor       =   &H00800000&
      Height          =   2295
      Left            =   15
      TabIndex        =   77
      Top             =   1830
      Visible         =   0   'False
      Width           =   3135
      Begin VB.ListBox multiven 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1950
         Left            =   60
         Style           =   1  'Checkbox
         TabIndex        =   78
         Top             =   195
         Visible         =   0   'False
         Width           =   2895
      End
   End
   Begin VB.Frame framoneda 
      Caption         =   "Moneda :"
      Height          =   615
      Left            =   7320
      TabIndex        =   35
      Top             =   5280
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
         Height          =   315
         ItemData        =   "frmRepoCrystal.frx":0442
         Left            =   120
         List            =   "frmRepoCrystal.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fraclipro 
      Height          =   855
      Left            =   240
      TabIndex        =   23
      Top             =   8040
      Visible         =   0   'False
      Width           =   1815
      Begin VB.ComboBox cmbclipro 
         Height          =   315
         ItemData        =   "frmRepoCrystal.frx":0461
         Left            =   120
         List            =   "frmRepoCrystal.frx":046B
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblclipro 
         AutoSize        =   -1  'True
         Caption         =   "Cliente / Proveedor"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1380
      End
   End
   Begin VB.CommandButton cmdCerrar 
      BackColor       =   &H00FFFFFF&
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
      Height          =   810
      Left            =   10440
      Picture         =   "frmRepoCrystal.frx":048B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   1080
   End
   Begin VB.CommandButton Pantalla 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Re&portar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   10440
      Picture         =   "frmRepoCrystal.frx":0D01
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1080
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   0
      Top             =   8040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame fracodclie 
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   5760
      TabIndex        =   27
      Top             =   1080
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox txt_cli 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   120
         MaxLength       =   8
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblCliente 
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
         Left            =   1920
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   2565
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fratipdoc 
      Caption         =   "Tipo de Documentos"
      ForeColor       =   &H00800000&
      Height          =   1575
      Left            =   3720
      TabIndex        =   79
      Top             =   1800
      Visible         =   0   'False
      Width           =   1815
      Begin VB.ListBox TIPDOC 
         Height          =   1185
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   81
         Top             =   240
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.ListBox SITUACION 
         Height          =   1185
         ItemData        =   "frmRepoCrystal.frx":1B7B
         Left            =   960
         List            =   "frmRepoCrystal.frx":1B7D
         Style           =   1  'Checkbox
         TabIndex        =   80
         Top             =   240
         Visible         =   0   'False
         Width           =   780
      End
   End
   Begin VB.Frame frazonas 
      Caption         =   "Filtro para Clientes :"
      Height          =   2610
      Left            =   15
      TabIndex        =   9
      Top             =   1905
      Visible         =   0   'False
      Width           =   10215
      Begin VB.CheckBox cheestado 
         Caption         =   "Mostrar Desactivos"
         Height          =   255
         Left            =   5040
         TabIndex        =   37
         Top             =   2280
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ListBox zonas 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1680
         Left            =   5640
         Style           =   1  'Checkbox
         TabIndex        =   13
         Top             =   240
         Width           =   4455
      End
      Begin VB.OptionButton opzonas 
         Caption         =   "Distrito"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   3360
         TabIndex        =   12
         Top             =   375
         Width           =   1815
      End
      Begin VB.OptionButton opzonas 
         Caption         =   "Provincia"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   3360
         TabIndex        =   11
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton opzonas 
         Caption         =   "Zonas"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   2
         Left            =   3360
         TabIndex        =   10
         Top             =   1320
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.Label lblzonas 
         AutoSize        =   -1  'True
         Caption         =   "Zonas :"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   4080
         TabIndex        =   14
         Top             =   600
         Width           =   1500
      End
   End
   Begin VB.Frame fracompra 
      Caption         =   "Opciones de Reg. de Compra: "
      ForeColor       =   &H00800000&
      Height          =   2895
      Left            =   15
      TabIndex        =   43
      Top             =   1800
      Visible         =   0   'False
      Width           =   10215
      Begin VB.CheckBox chenc 
         Caption         =   "Solo Notas de Credito"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CheckBox checompras 
         Caption         =   "Solo Compras de Mercaderia"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2280
         TabIndex        =   56
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox moneda 
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   54
         Text            =   "T"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox codsunat 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   51
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox difigv 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   4680
         TabIndex        =   50
         Top             =   1920
         Width           =   855
      End
      Begin VB.CheckBox cheigv 
         Caption         =   "Consistenciar el Impto.: Por  Diferencia (+/-) >= : "
         ForeColor       =   &H00000080&
         Height          =   615
         Left            =   2880
         TabIndex        =   49
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtorden 
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   5760
         MaxLength       =   1
         TabIndex        =   47
         Text            =   "F"
         Top             =   2280
         Width           =   375
      End
      Begin VB.OptionButton opcompra 
         Caption         =   "Por Gastos"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   46
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton opcompra 
         Caption         =   "Por Proveedor "
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   45
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton opcompra 
         Caption         =   "Todo el Registro"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Moneda: (S/D/A/T):"
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   55
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label nomsunat 
         Height          =   375
         Left            =   2280
         TabIndex        =   53
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Por Tipo de Doc. :"
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Ordenado por: (F/D/R) "
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   3960
         TabIndex        =   48
         Top             =   2280
         Width           =   1815
      End
   End
   Begin VB.Frame fraArti 
      Caption         =   "Filtro de Articulos :"
      ForeColor       =   &H00800000&
      Height          =   5055
      Left            =   15
      TabIndex        =   5
      Top             =   1755
      Visible         =   0   'False
      Width           =   10215
      Begin VB.ComboBox famix 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   105
         Top             =   840
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.ComboBox art_subgru 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   104
         Top             =   2040
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.ComboBox art_numero 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   2640
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.ComboBox LINEAS 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   3240
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.ComboBox art_marca 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   101
         Top             =   3840
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.ComboBox art_plancha 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   100
         Top             =   4440
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.ComboBox subfami 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   99
         Top             =   1440
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox i_codart2 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
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
         Height          =   1410
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   2385
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label lblporcentaje 
         Caption         =   "Porcentaje"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   225
         TabIndex        =   113
         Top             =   3840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblarti 
         Caption         =   "Codigo de Articulo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   7
         Left            =   3360
         TabIndex        =   94
         Top             =   240
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label lblarti 
         Caption         =   "Lote de Articulo :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   6
         Left            =   3480
         TabIndex        =   87
         Top             =   4200
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label lblarti 
         Caption         =   "Marca de Articulo :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   5
         Left            =   3480
         TabIndex        =   86
         Top             =   3600
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label lblarti 
         Caption         =   "Sub Gupo de Articulos :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   3480
         TabIndex        =   85
         Top             =   3000
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label lblarti 
         Caption         =   "Clase de Articulos :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   84
         Top             =   2400
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label lblarti 
         Caption         =   "Linea por Articulo :"
         BeginProperty Font 
            Name            =   "Arial"
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
         Left            =   3480
         TabIndex        =   83
         Top             =   1800
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label lblarti 
         Caption         =   "Sub-División de Articulos :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   82
         Top             =   1200
         Visible         =   0   'False
         Width           =   3015
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
         Height          =   315
         Left            =   6360
         TabIndex        =   21
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lblarti 
         Caption         =   "Divisiones de Articulos :"
         BeginProperty Font 
            Name            =   "Arial"
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
         Left            =   3480
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   3015
      End
   End
   Begin VB.Label lblusers 
      Caption         =   "Usuarios"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   465
      TabIndex        =   111
      Top             =   6060
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblnom 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Dia de Visita :"
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
      TabIndex        =   109
      Tag             =   "27"
      Top             =   5685
      Visible         =   0   'False
      WhatsThisHelpID =   17
      Width           =   975
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
      Left            =   9120
      TabIndex        =   30
      Top             =   8160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblproceso 
      BackColor       =   &H008B4914&
      Caption         =   "Procesando. . ."
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
      Height          =   1095
      Left            =   10440
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4914&
      BorderStyle     =   1  'Fixed Single
      Height          =   7095
      Index           =   5
      Left            =   10335
      TabIndex        =   70
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "RCRYSTAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WFAMI2 As Integer
Dim NIVEL_MAX As Integer
Dim LOC_RUC As String
Dim Wfile As String
Dim xl As Object
Dim wran1 As String
Dim wran2 As String
Dim wranF As String
Dim ART_CLASES As String
Dim ART_ARTICULO As String
Dim FAR_FECHAS As String
Dim ART_LINEAS As String

Dim f1 As Integer
Dim c1 As Integer
Dim WFORM As String
Dim REP_FECHA1 As String
Dim REP_FECHA2 As String
Dim VAR_ACTIVAR As Integer
Dim WCOD_ORIGINAL As Currency
Dim loc_key As Integer
Dim loc_cp As String * 1
Dim WW_CODVEN As Integer
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset
Dim PS_REP02 As rdoQuery
Dim llave_rep02 As rdoResultset
Dim PS_REP03 As rdoQuery
Dim llave_rep03 As rdoResultset
Dim PS_REP04 As rdoQuery
Dim llave_rep04 As rdoResultset


Private Sub art_numero_Click()
WFAMI2 = val(Trim(Right(art_numero.Text, 6)))
PUB_TIPREG = 131
LLENADO_SUBFAM LINEAS, WFAMI2
art_marca.Clear


End Sub

Private Sub art_subgru_Click()
WFAMI2 = val(Trim(Right(art_subgru.Text, 6)))
PUB_TIPREG = 130
LLENADO_SUBFAM art_numero, WFAMI2
LINEAS.Clear
art_marca.Clear



End Sub

Private Sub cmdCerrar_Click()
Unload RCRYSTAL
End Sub

Private Sub codsunat_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii <> 13 Then Exit Sub
If Trim(codsunat.Text) = "" Then Exit Sub
PUB_TIPREG = 50
PUB_NUMTAB = val(codsunat.Text)
PUB_CODCIA = "00"
SQ_OPER = 1
LEER_TAB_LLAVE
If tab_llave.EOF Then
  MsgBox "No Existe Codigo de Sunat ", 48, Pub_Titulo
  Azul codsunat, codsunat
  Exit Sub
End If
nomsunat.Caption = tab_llave!tab_NOMLARGO
If txtorden.Visible Then txtorden.SetFocus

End Sub

Private Sub cta1_Change()
ncta1.Caption = ""
End Sub

Private Sub cta1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 cta1.Text = ""
 Exit Sub
End If

If KeyAscii <> 13 Then Exit Sub
If KeyAscii = 13 Then
     If Trim(Left(cta1.Text, 1)) = "*" Then
       BUSCAR_CTA 1, cta1
       Exit Sub
     End If
End If

If Trim(cta1.Text) = "" Then Exit Sub
    SQ_OPER = 1
    PUB_CUENTA = Trim(cta1.Text)
    LEER_COM_LLAVE
    If com_llave.EOF Then
     MsgBox "Cuanta No Existe...", 48, Pub_Titulo
     Azul cta1, cta1
     Exit Sub
    End If
    If com_llave!COM_NIVEL <> NIVEL_MAX Then
      MsgBox "No Procede.. Cuanta no es Analitica...", 48, Pub_Titulo
      Azul cta1, cta1
      Exit Sub
    End If
    ncta1.Caption = Trim(com_llave!com_descripcion)
    Azul cta2, cta2
End Sub

Private Sub cta2_Change()
ncta2.Caption = ""
End Sub

Private Sub cta2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 cta2.Text = ""
 Exit Sub
End If
If KeyAscii <> 13 Then Exit Sub
If KeyAscii = 13 Then
     If Trim(Left(cta2.Text, 1)) = "*" Then
       BUSCAR_CTA 1, cta2
       Exit Sub
     End If
End If


If Trim(cta2.Text) = "" Then Exit Sub
    SQ_OPER = 1
    PUB_CUENTA = Trim(cta2.Text)
    LEER_COM_LLAVE
    If com_llave.EOF Then
     MsgBox "Cuanta No Existe...", 48, Pub_Titulo
     Azul cta2, cta2
     Exit Sub
    End If
    If com_llave!COM_NIVEL <> NIVEL_MAX Then
      MsgBox "No Procede.. Cuanta no es Analitica...", 48, Pub_Titulo
      Azul cta2, cta2
      Exit Sub
    End If
    ncta2.Caption = Trim(com_llave!com_descripcion)
    Azul cta3, cta3
End Sub

Private Sub cta3_Change()
ncta3.Caption = ""
End Sub

Private Sub cta3_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 cta3.Text = ""
 Exit Sub
End If

If KeyAscii <> 13 Then Exit Sub
If KeyAscii = 13 Then
     If Trim(Left(cta3.Text, 1)) = "*" Then
       BUSCAR_CTA 1, cta3
       Exit Sub
     End If
End If

If Trim(cta3.Text) = "" Then Exit Sub
    SQ_OPER = 1
    PUB_CUENTA = Trim(cta3.Text)
    LEER_COM_LLAVE
    If com_llave.EOF Then
     MsgBox "Cuanta No Existe...", 48, Pub_Titulo
     Azul cta3, cta3
     Exit Sub
    End If
    If com_llave!COM_NIVEL <> NIVEL_MAX Then
      MsgBox "No Procede.. Cuanta no es Analitica...", 48, Pub_Titulo
      Azul cta3, cta3
      Exit Sub
    End If
    ncta3.Caption = Trim(com_llave!com_descripcion)
    If Pantalla.Enabled Then Pantalla.SetFocus

End Sub

Private Sub cheigv_Click()
If cheigv.Value = 1 Then
difigv.SetFocus
Else
codsunat.SetFocus
End If
End Sub

Private Sub difigv_KeyPress(KeyAscii As Integer)
  SOLO_DECIMAL difigv, KeyAscii
  If KeyAscii <> 13 Then Exit Sub
  codsunat.SetFocus
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
WFAMI2 = val(Trim(Right(famix.Text, 6)))
PUB_TIPREG = 123
LLENADO_SUBFAM subfami, WFAMI2

art_subgru.Clear
art_numero.Clear
LINEAS.Clear
art_marca.Clear

On Error GoTo sigue
subfami.ListIndex = wpos
Exit Sub
sigue:
Resume Next

End Sub

Private Sub Form_Activate()
lblarti(0).Caption = "División"
lblarti(1).Caption = "Familia"
lblarti(2).Caption = "Linea"
lblarti(3).Caption = "Sub Linea"
lblarti(4).Caption = "Marca"
lblarti(5).Caption = "Clase"
lblarti(6).Caption = "Lote"

End Sub

Private Sub Form_Load()
Dim xcuenta As Integer
If Not cop_llave.EOF Then
For fila = 1 To 6
  If cop_llave.rdoColumns(fila) <> 0 Then
     'wCOM_NIVEL(i) = cop_llave.rdoColumns(i)
     NIVEL_MAX = fila
  End If
Next fila
End If
lblarti(0).Caption = "División"
lblarti(1).Caption = "Familia"
lblarti(2).Caption = "Linea"
lblarti(3).Caption = "Sub Linea"
lblarti(4).Caption = "Marca"
lblarti(5).Caption = "Clase"
lblarti(6).Caption = "Lote"

   
   
LOC_RUC = ""
VAR_ACTIVAR = 0
CentrarFormulario MDIForm1, Me
Screen.MousePointer = 11
If retra_llave.EOF Then
   Screen.MousePointer = 0
   Exit Sub
End If
Screen.MousePointer = 0
Wfile = Trim(retra_llave(3))
WFORM = Trim(retra_llave(7))
LBLRUTA.Caption = Wfile

pub_cadena = "SELECT PAR_NOMBRE FROM PARGEN WHERE PAR_CODCIA = ? ORDER BY PAR_NOMBRE "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
If Wfile = "CLIVENTA.RPT" Or Wfile = "CLILINVEN.RPT" Then
  rango.Visible = True
End If
If Wfile = "CLIVTASPOR.RPT" Then
  txtporcentaje.Visible = True
  lblporcentaje.Visible = True
End If
If retra_llave!TRA_S14 = 1 Or retra_llave!TRA_S15 = 1 Or retra_llave!TRA_C1 = 1 Then
    cmdmoneda.ListIndex = 0
    framoneda.Visible = True
End If
If retra_llave!tra_c2 = 1 Then
  i_codart2.TabIndex = 0
  fraArti.Visible = True
  lblarti(6).Visible = True
  lblarti(7).Visible = True
  i_codart2.Visible = True
End If


If retra_llave!tra_con6 = 1 Then
  i_codart2.TabIndex = 0
  fraArti.Visible = True
  lblarti(6).Visible = True
  lblarti(7).Visible = True
  i_codart2.Visible = True
ElseIf retra_llave!tra_con7 = 1 Then
 PUB_CODCIA = LK_CODCIA
 If LK_EMP_PTO = "A" Then
      PUB_CODCIA = "00"
 End If
 LLENADOS_LISTAS fami, 122
 fami.TabIndex = 0
 lblarti(0).Visible = True
 subfami.TabIndex = 1
 fraArti.Visible = True
 fami.Visible = True
ElseIf retra_llave!tra_con9 = 1 Then
 PUB_CODCIA = LK_CODCIA
 If LK_EMP_PTO = "A" Then
      PUB_CODCIA = "00"
 End If
 LLENADOS famix, 122
 famix.TabIndex = 0
 fami.Visible = False
 'lblarti(0).Visible = True
 'lblarti(1).Visible = True
 subfami.TabIndex = 1
 subfami.Visible = True
 lblarti(1).Visible = True
 lblarti(0).Visible = True
 famix.Visible = True
 fraArti.Visible = True
End If
If retra_llave!TRA_ACT12 = 1 Then
 PUB_CODCIA = LK_CODCIA
 LLENADOS art_numero, 130
 art_numero.Visible = True
 lblarti(4).Visible = True
 fraArti.Visible = True
End If
If retra_llave!TRA_ACT13 = 1 Then
 PUB_CODCIA = LK_CODCIA
 LLENADOS art_marca, 133
 art_marca.Visible = True
 lblarti(5).Visible = True
 fraArti.Visible = True
End If
If retra_llave!TRA_S5 = 1 Then
  LLENADOS art_plancha, 132
  lblarti(6).Visible = True
  art_plancha.Visible = True
End If

If retra_llave!tra_s12 = 1 Or retra_llave!TRA_S9 = 1 Or retra_llave!tra_ACT11 = 1 Or retra_llave!TRA_S6 = 1 Or Wfile = "KARDEX_CLASES" Then
 LISCIA.Visible = True
 LISCIA.Clear
 xcuenta = 0
 For fila = 1 To 30 Step 2
   PUB_CODCIA = Mid(Trim(par_llave!par_art_cias), fila, 2)
   If Trim(PUB_CODCIA) = "" Then Exit For
   xcuenta = xcuenta + 1
   PS_REP02(0) = PUB_CODCIA
   llave_rep02.Requery
   LISCIA.AddItem PUB_CODCIA & " - " & Trim(llave_rep02!PAR_NOMBRE)
 Next fila
 If LISCIA.ListCount = 0 Then
   LISCIA.Visible = False
   lblciaact.Caption = LK_CODCIA & "-" & Trim(par_llave!PAR_NOMBRE)
 End If
 
 For fila = 0 To LISCIA.ListCount - 1
  LISCIA.ListIndex = fila
  If Left(LISCIA.Text, 2) = LK_CODCIA Then LISCIA.Selected(fila) = True
 Next fila
'End If

End If

If retra_llave!TRA_ACT15 = 1 Then
 PUB_CODCIA = LK_CODCIA
 If LK_EMP_PTO = "A" Then
      PUB_CODCIA = "00"
 End If
 LLENADOS LINEAS, 131
 LINEAS.TabIndex = 0
 LINEAS.Visible = True
 lblarti(3).Visible = True
 fraArti.Visible = True
End If
If retra_llave!TRA_S2 = 1 Then
 PUB_CODCIA = LK_CODCIA
 If LK_EMP_PTO = "A" Then
      PUB_CODCIA = "00"
 End If
 LLENADOS art_subgru, 129
 art_subgru.TabIndex = 0
 art_subgru.Visible = True
 lblarti(2).Visible = True
 fraArti.Visible = True
End If


If retra_llave!TRA_S10 = 1 Then
   cheflag.Caption = retra_llave!TRA_l3
   fraflag.Visible = True
End If
If retra_llave!TRA_GRU2 = 1 Then
   CHEFLAG2.Caption = retra_llave!TRA_l4
   FRAFLAG2.Visible = True
End If

If retra_llave!TRA_CON8 = 1 Then
 PUB_CODCIA = "00"
 LLENADOS_LISTAS zonas, 35
 frazonas.Visible = True
 opzonas(0).Caption = BUSCA_ETIQUETA(10)
 opzonas(1).Caption = BUSCA_ETIQUETA(11)
 opzonas(2).Caption = BUSCA_ETIQUETA(12)
End If
If retra_llave!TRA_CON4 = 1 Then
  fracodclie.Caption = "Cliente "
  fracodclie.Visible = True
  txt_cli.Visible = True
  txt_cli.TabIndex = 0
  lblCliente.Visible = True
  loc_cp = "C"
End If
If retra_llave!TRA_CON5 = 1 Then
  txt_cli.TabIndex = 0
  fracodclie.Caption = "Proveedor "
  fracodclie.Visible = True
  txt_cli.Visible = True
  lblCliente.Visible = True
  loc_cp = "P"
End If
If retra_llave!TRA_CON11 = 1 Then
  fraclipro.Visible = True
  cmbclipro.ListIndex = 0
End If
If retra_llave!tra_GRU1 = 1 Or retra_llave!TRA_S11 = 1 Or retra_llave!TRA_S8 = 1 Or retra_llave!TRA_ACT5 = 1 Or retra_llave!TRA_CON14 = 1 Or retra_llave!tra_con1 = 1 Or retra_llave!tra_con10 = 1 Or retra_llave!tra_act8 = 1 Or retra_llave!tra_con12 = 1 Then
 frafechas.Visible = True
 If retra_llave!tra_con10 = 1 Then
 lblcampo1.Caption = "Fec. Vcto. : "
 ElseIf retra_llave!tra_act8 = 1 Then
 lblcampo1.Caption = "Fec. Ingreso. : "
 ElseIf retra_llave!TRA_S8 = 1 Then
 lblcampo1.Caption = "Fechas Contable: "
 ElseIf retra_llave!tra_GRU1 = 1 Then
  lblcampo1.Caption = "Fec. Proc.: "
 ElseIf retra_llave!tra_con1 = 1 Then
  lblcampo1.Caption = "Fec. Emis.: "
 Else
 lblcampo1.Caption = "Fecha de Inicial : "
 End If
 txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
' txtCampo1.Text = "01/01/2001"
 txtCampo1.Mask = "##/##/####"
 If retra_llave!TRA_S8 = 1 Then
  lblcampo2.Caption = " "
 Else
 lblcampo2.Caption = "Fecha de Final: "
 End If
 txtCampo2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
' txtCampo2.Text = "31/01/2001"
 txtCampo2.Mask = "##/##/####"
 txtCampo1.TabIndex = 0
 txtCampo2.TabIndex = 1
 
End If

 If retra_llave("tra_key") = 9304 Or retra_llave("tra_key") = 9118 Or retra_llave("tra_key") = 9311 Then   'agregado por mic
    PUB_CODCIA = LK_CODCIA
    LLENADOS cboDias, 17
    lblnom.Visible = True
    cboDias.Visible = True
 ElseIf retra_llave("tra_key") = 9302 Or retra_llave("tra_key") = 9303 Or retra_llave("tra_key") = 8121 Or retra_llave("tra_key") = 8716 Or retra_llave("tra_key") = 8116 Or retra_llave("tra_key") = 8132 Or retra_llave("tra_key") = 8133 Or retra_llave("tra_key") = 8134 Or retra_llave("tra_key") = 8510 Or retra_llave("tra_key") = 8522 Or retra_llave("tra_key") = 8531 Or retra_llave("tra_key") = 8716 Then 'agregado por gts para filtrar usuarios en reporte de caja
    cboUsers.Visible = True
    lblusers.Visible = True
    cboUsers.AddItem ""
    usu.MoveFirst
    Do While Not usu.EOF
        cboUsers.AddItem Trim(usu("USU_KEY"))
        usu.MoveNext
    Loop
 End If
'txtCampo1.Text = "01/01/2001"
'txtCampo2.Text = "31/01/2001"


If retra_llave!TRA_CON3 = 1 Then
  LLENA_VENDEDORES
  fraven.Visible = True
  multiven.Visible = True
  multiven.TabIndex = 1
End If
If retra_llave!TRA_CON2 = 1 Then
 Frame2.Visible = True
End If
If retra_llave!TRA_ACT6 = 1 Then
  framoneda.Visible = True
  cmdmoneda.ListIndex = 0
End If
If retra_llave!TRA_ACT7 = 1 Then
  cheestado.Visible = True
End If
If retra_llave!TRA_S13 = 1 Then
  fratipo.Visible = True
    PUB_TIPREG = 230
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 2
    LEER_TAB_LLAVE
    lsttipo.ToolTipText = "TAB_TIPREG = 230"
    lsttipo.Clear
    Do Until tab_mayor.EOF
        lsttipo.AddItem tab_mayor!tab_NOMLARGO & String(80, " ") & tab_mayor!TAB_NUMTAB
        If tab_mayor!TAB_NUMTAB = 1 Then lsttipo.Selected(tab_mayor.AbsolutePosition - 1) = True
        tab_mayor.MoveNext
    Loop
End If

If retra_llave!TRA_S7 = 1 Then
    PUB_TIPREG = 2
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 2
    LEER_TAB_LLAVE
    listacal.ToolTipText = "TAB_TIPREG = 2"
    listacal.Clear
    Do Until tab_mayor.EOF
        listacal.AddItem tab_mayor!tab_NOMLARGO & String(80, " ") & tab_mayor!TAB_NUMTAB
        If tab_mayor!TAB_NUMTAB = 1 Then listacal.Selected(tab_mayor.AbsolutePosition - 1) = True
        tab_mayor.MoveNext
    Loop
    fracal.Visible = True
End If

lblformulas.Caption = ""
If retra_llave!tra_ACT1 = 1 Then lblformulas.Caption = lblformulas.Caption + "; CIA   "
If retra_llave!tra_ACT2 = 1 Then lblformulas.Caption = lblformulas.Caption + "; DIA   "
If retra_llave!tra_ACT10 = 1 Then lblformulas.Caption = lblformulas.Caption + "; FECHAS   "
If retra_llave!TRA_ACT4 = 1 Then
 LBLTIPDOC.Visible = True
 SITUACION.Visible = True
 fratipdoc.Visible = True
 TIPDOC.Visible = True
 PUB_CODCIA = "00"
 LLENADOS_LISTAS SITUACION, 133
 LLENADOS_LISTAS TIPDOC, 8
End If

lblreporte.Caption = Trim(retra_llave(1))
If Wfile = "SALINI.RPT" Then
  txtCampo1.Enabled = False
  txtCampo2.Enabled = False
  Pantalla.TabIndex = 0
End If
If Wfile = "REG_BANCOS" Then
  frapasa.Visible = True
  If Not cop_llave.EOF Then
   cop_fecha1.Caption = Format(cop_llave!cop_fecha_proceso, "dd/mm/yy")
   cop_fecha2.Caption = Format(cop_llave!cop_fecha_proceso2, "dd/mm/yy")
  End If
  frag.Visible = True
End If
If Wfile = "REG_COMPRA_COM" Then
  frapasa.Visible = True
  If Not cop_llave.EOF Then
   cop_fecha1.Caption = Format(cop_llave!cop_fecha_proceso, "dd/mm/yy")
   cop_fecha2.Caption = Format(cop_llave!cop_fecha_proceso2, "dd/mm/yy")
  End If
  frag.Visible = True
  fracompra.Visible = True
  fracodclie.Visible = False
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
  If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
  GoTo posicion
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo posicion
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
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
 PUB_KEY = val(i_codart2.Text)
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
    i_nomarti.Caption = Trim(art_LLAVE!art_nombre)
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
     i_nomarti.Caption = Trim(art_llave_alt!art_nombre)
     ListView1.Visible = False
     If Pantalla.Enabled Then Pantalla.SetFocus
     Exit Sub
  Else
    If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
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
       PUB_KEY = val(ListView1.ListItems.Item(loc_key).SubItems(1))
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
      i_nomarti.Caption = Trim(art_LLAVE!art_nombre)
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
Dim VAR
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
    VAR = Asc(i_codart2.Text)
    VAR = VAR + 1
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      numarchi = 3
      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND  (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND ART_CODCIA = '" & LK_CODCIA & "' AND ART_ALTERNO BETWEEN '" & i_codart2.Text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
    Else
    
    numarchi = 0
        If VAR_ACTIVAR <> 666 Then
            ''archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND  ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
              archi = "SELECT ARTI.ART_KEY,ARTI.ART_CODCIA, ARTI.ART_NOMBRE, ARTI.ART_ALTERNO, ARTICULO.ARM_STOCK, PRECIOS.PRE_EQUIV, TABLAS.TAB_NOMLARGO AS DIVI, TABLAS_1.TAB_NOMLARGO AS LINEA, PRECIOS.PRE_PRE1, PRECIOS.PRE_PRE2,ARTI.ART_CUENTA_CONTAB,ARTI.ART_CUENTA_CONTAB_C  "
              archi = archi & "FROM ARTI INNER JOIN ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA INNER JOIN PRECIOS ON ARTI.ART_KEY = PRECIOS.PRE_CODART AND ARTI.ART_CODCIA = PRECIOS.PRE_CODCIA INNER JOIN TABLAS ON ARTI.ART_CODCIA = TABLAS.TAB_CODCIA AND ARTI.ART_FAMILIA = TABLAS.TAB_NUMTAB INNER JOIN TABLAS TABLAS_1 ON ARTI.ART_CODCIA = TABLAS_1.TAB_CODCIA AND ARTI.ART_SUBFAM = TABLAS_1.TAB_NUMTAB "
              archi = archi & "WHERE (TABLAS.TAB_TIPREG = 122) AND (TABLAS_1.TAB_TIPREG = 123) AND (PRECIOS.PRE_FLAG_UNIDAD = 'A') AND ARTI.ART_CODCIA = '" & LK_CODCIA & "' AND ARTI.ART_CALIDAD = 1 AND (ARTI.ART_FLAG_STOCK = 'M' OR ARTI.ART_FLAG_STOCK = 'P' OR ARTI.ART_FLAG_STOCK = 'C') AND ARTI.ART_NOMBRE BETWEEN '" & Trim(i_codart2.Text) & "%' AND  '" & VAR & "' ORDER BY ARTI.ART_NOMBRE"
        Else
            If KeyCode = 13 Then
              archi = "SELECT ARTI.ART_KEY,ARTI.ART_CODCIA, ARTI.ART_NOMBRE, ARTI.ART_ALTERNO, ARTICULO.ARM_STOCK, PRECIOS.PRE_EQUIV, TABLAS.TAB_NOMLARGO AS DIVI, TABLAS_1.TAB_NOMLARGO AS LINEA, PRECIOS.PRE_PRE1, PRECIOS.PRE_PRE2,ARTI.ART_CUENTA_CONTAB,ARTI.ART_CUENTA_CONTAB_C  "
              archi = archi & "FROM ARTI INNER JOIN ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA INNER JOIN PRECIOS ON ARTI.ART_KEY = PRECIOS.PRE_CODART AND ARTI.ART_CODCIA = PRECIOS.PRE_CODCIA INNER JOIN TABLAS ON ARTI.ART_CODCIA = TABLAS.TAB_CODCIA AND ARTI.ART_FAMILIA = TABLAS.TAB_NUMTAB INNER JOIN TABLAS TABLAS_1 ON ARTI.ART_CODCIA = TABLAS_1.TAB_CODCIA AND ARTI.ART_SUBFAM = TABLAS_1.TAB_NUMTAB "
              archi = archi & "WHERE (TABLAS.TAB_TIPREG = 122) AND (TABLAS_1.TAB_TIPREG = 123) AND (PRECIOS.PRE_FLAG_UNIDAD = 'A') AND ARTI.ART_CODCIA = '" & LK_CODCIA & "' AND ARTI.ART_CALIDAD = 1 AND (ARTI.ART_FLAG_STOCK = 'M' OR ARTI.ART_FLAG_STOCK = 'P' OR ARTI.ART_FLAG_STOCK = 'C') AND ARTI.ART_NOMBRE like '" & Trim(i_codart2.Text) & "%' ORDER BY ARTI.ART_NOMBRE"
            Else
                Exit Sub
            End If
        End If
    
    
      'numarchi = 0
      ''archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND ART_CODCIA = '" & LK_CODCIA & "' AND ART_NOMBRE BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
     ' archi = "SELECT ARTI.ART_KEY,ARTI.ART_CODCIA, ARTI.ART_NOMBRE, ARTI.ART_ALTERNO, ARTICULO.ARM_STOCK, PRECIOS.PRE_EQUIV, TABLAS.TAB_NOMLARGO AS DIVI, TABLAS_1.TAB_NOMLARGO AS LINEA "
    'archi = archi & "FROM ARTI INNER JOIN ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA INNER JOIN PRECIOS ON ARTI.ART_KEY = PRECIOS.PRE_CODART AND ARTI.ART_CODCIA = PRECIOS.PRE_CODCIA INNER JOIN TABLAS ON ARTI.ART_CODCIA = TABLAS.TAB_CODCIA AND ARTI.ART_FAMILIA = TABLAS.TAB_NUMTAB INNER JOIN TABLAS TABLAS_1 ON ARTI.ART_CODCIA = TABLAS_1.TAB_CODCIA AND ARTI.ART_SUBFAM = TABLAS_1.TAB_NUMTAB "
    'archi = archi & "WHERE (TABLAS.TAB_TIPREG = 122) AND (TABLAS_1.TAB_TIPREG = 123) AND (PRECIOS.PRE_FLAG_UNIDAD = 'A') AND ARTI.ART_CODCIA = '" & LK_CODCIA & "' AND ARTI.ART_CALIDAD = 1 AND ARTI.ART_FLAG_STOCK = 'M' AND ARTI.ART_NOMBRE BETWEEN '" & Trim(i_codart2.Text) & "%' AND  '" & var & "' ORDER BY ARTI.ART_NOMBRE"
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
   If loc_key + 8 > ListView1.ListItems.count Then
      ListView1.ListItems.Item(ListView1.ListItems.count).EnsureVisible
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

Private Sub LINEAS_Click()
WFAMI2 = val(Trim(Right(LINEAS.Text, 6)))
PUB_TIPREG = 132
LLENADO_SUBFAM art_marca, WFAMI2

End Sub

Private Sub ListView1_DblClick()
 loc_key = ListView1.SelectedItem.index
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
 loc_key = ListView1.SelectedItem.index
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
 loc_key = ListView2.SelectedItem.index
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
 loc_key = ListView2.SelectedItem.index
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



Private Sub moneda_Change()
moneda.Text = UCase(moneda.Text)
End Sub

Private Sub moneda_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Trim(moneda.Text) = "S" Or Trim(moneda.Text) = "D" Or Trim(moneda.Text) = "A" Or Trim(moneda.Text) = "T" Then
Else
  MsgBox "No es parametro..verificar", 48, Pub_Titulo
  Azul moneda, moneda
  Exit Sub
End If
moneda.Text = UCase(moneda.Text)
End Sub

Private Sub multiven_Click()
If Pantalla.Enabled = True Then
If LK_EMP = "PAR" Then
 WW_CODVEN = val(Left(multiven.Text, 3))
 PUB_CODCIA = "00"
 LLENADOS_LISTAS zonas, 35
 End If
End If
End Sub

Private Sub op1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  op2.SetFocus
End If
End Sub

Private Sub opcompra_Click(index As Integer)
If index = 0 Then   ' todo
frag.Visible = False
fracodclie.Visible = False
txt_cli.Text = ""
cta1.Text = ""
cta2.Text = ""
cta3.Text = ""
'checompras.Visible = True
'checompras.Value = 0
If Pantalla.Enabled Then Pantalla.SetFocus
End If
If index = 1 Then   ' Proveedor
'checompras.Visible = False
'checompras.Value = 0
fracodclie.Visible = True
frag.Visible = False
cta1.Text = ""
cta2.Text = ""
cta3.Text = ""
txt_cli.SetFocus
End If
If index = 2 Then   ' Gastos
'checompras.Visible = False
checompras.Value = 0
frag.Visible = True
fracodclie.Visible = False
txt_cli.Text = ""
cta1.SetFocus
End If


End Sub

Private Sub opzonas_Click(index As Integer)
Dim cod As Integer
lblzonas.Caption = Trim(opzonas(index).Caption) & " :"
If index = 0 Then
  cod = 20
ElseIf index = 1 Then
  cod = 30
ElseIf index = 2 Then
  cod = 35
End If
PUB_CODCIA = "00"
LLENADOS_LISTAS zonas, cod
zonas.SetFocus

End Sub

Private Sub Pantalla_Click()
If Wfile = "KARDEX_CLASES" Then
   KARDEX_CLASES
ElseIf Wfile = "RESU_KARDEX" Then
   RESU_KARDEX
ElseIf Wfile = "REG_COMPRA_COM" Then
   REG_COMPRA_COM
ElseIf Wfile = "REG_BANCOS" Then
   REG_BANCOS
ElseIf Wfile = "MOVI_BANCOS" Then
'   CHEQUEO_DESCTO
   MOVI_BANCO
Else
   PRO_REPORTE (0)
End If
ART_CLASES = ""
ART_ARTICULO = ""
End Sub
Public Sub LLENADOS_LISTAS(cont As ListBox, tip As Integer)
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
       If PUB_TIPREG = 35 And LK_EMP = "PAR" Then
          If val(tab_mayor!TAB_CODART) = WW_CODVEN Then
            cont.AddItem tab_mayor!tab_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
          End If
       Else
           cont.AddItem tab_mayor!tab_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
       End If
       tab_mayor.MoveNext
    Loop
End Sub

Public Sub LLENADOS(cont As ComboBox, tip As Integer)
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
       If PUB_TIPREG = 35 And LK_EMP = "PAR" Then
          If val(tab_mayor!TAB_CODART) = WW_CODVEN Then
            cont.AddItem tab_mayor!tab_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
          End If
       Else
           cont.AddItem tab_mayor!tab_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
       End If
       tab_mayor.MoveNext
    Loop
End Sub
Public Sub LLENADO_SUBFAM(ctlCombo As ComboBox, ByVal wfami As Integer)
On Error GoTo SALE
Dim CONTA As Integer
    CONTA = -1
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       PUB_CODCIA = "00"
    End If
    PUB_CODART = wfami
    SQ_OPER = 3
    LEER_TAB_LLAVE
    Select Case ctlCombo.Name
      Case Is = "art_subfam"
       ctlCombo.ToolTipText = "TAB_TIPREG = 123"
      Case Is = "art_grupo"
       ctlCombo.ToolTipText = "TAB_TIPREG = 129"
      Case Is = "art_numero"
       ctlCombo.ToolTipText = "TAB_TIPREG = 130"
      Case Is = "art_linea"
       ctlCombo.ToolTipText = "TAB_TIPREG = 131"
      Case Is = "art_marca"
       ctlCombo.ToolTipText = "TAB_TIPREG = 132"
    End Select
    ctlCombo.Clear
    ctlCombo.AddItem "Todos"
    Do Until tab_menor.EOF
        ctlCombo.AddItem tab_menor!tab_NOMLARGO & String(80, " ") & Trim(CStr(tab_menor!TAB_NUMTAB))
        DoEvents
        CONTA = CONTA + 1
        tab_menor.MoveNext
    Loop
    If ctlCombo.ListCount > 0 Then ctlCombo.ListIndex = 0
Exit Sub
SALE:
Resume Next
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

Public Sub PRO_REPORTE(x As Integer)
Dim cade2 As String
Dim wsalmacenes As String
Dim xcuenta2 As Integer
Dim ENTRO As Integer
Dim wf1, wf2, wf3, wf4, wf5, wf6, wf7, wf8, wf9, wf10
Dim DIA, MES, ANO
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
If retra_llave!TRA_CON4 = "1" Or retra_llave!TRA_CON5 = "1" Then
'  If Trim(txt_cli.Text) = "" Then
'      MsgBox "Verificar Codigo ", 48, Pub_Titulo
'      Exit Sub
'  End If
End If

ART_LINEAS = ""
ART_CLASES = ""

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
wsalmacenes = ""
Pantalla.Enabled = False
cmdCerrar.Enabled = False

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
  If retra_llave!tra_rep1 = "1" Then
     Reportes.ReportFileName = PUB_RUTA_OTRO & "PTOVTA\" & Wfile
     wcodcia = LK_CODCIA
  Else
    Reportes.ReportFileName = PUB_RUTA_OTRO & Wfile
    wcodcia = LK_CODCIA
  End If
  Reportes.WindowTitle = "Reporte :  " & Trim(retra_llave(1)) & " - Archivo:(" & Wfile & ")"
  ProgBar.Value = 4
  Reportes.Destination = crptToWindow
  Reportes.WindowLeft = 2
  Reportes.WindowTop = 70
  Reportes.WindowWidth = 790
  Reportes.WindowHeight = 475
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
  Reportes.Formulas(80) = "" 'FLAG_D
  Reportes.Formulas(81) = ""  'FLAG_SL
  ProgBar.Value = 6
  pub_cadena = ""
  wmensa = ""
  If retra_llave!TRA_CON2 = 1 And val(txt_key.Text) <> 0 Then
    If pub_cadena = "" Then
       pub_cadena = "{CCMAEST.CCM_CODBAN} = " & Trim(txt_key.Text)
    Else
        pub_cadena = pub_cadena + " AND " + "{CCMAEST.CCM_CODBAN} = " & Trim(txt_key.Text)
    End If
    If retra_llave!TRA_S9 <> 1 Then
      If pub_cadena = "" Then
           pub_cadena = pub_cadena + " {CCMAEST.CCM_CODCIA} = " + "'" & par_llave!PAR_CIACCM & "' "
      Else
           pub_cadena = pub_cadena + " AND " + " {CCMAEST.CCM_CODCIA} = " + "'" & par_llave!PAR_CIACCM & "' "
      End If
    End If
  End If
  If retra_llave!TRA_S1 = 1 Then
    If pub_cadena = "" Then
       pub_cadena = "{COMAEST.COM_CODCIA} = '" & LK_CODCIA & "'"
    Else
        pub_cadena = pub_cadena + " AND " + "{COMAEST.COM_CODCIA} = '" & LK_CODCIA & "'"
    End If
  End If
  

  If (retra_llave!TRA_CON4 = 1 Or retra_llave!TRA_CON5 = 1) And val(txt_cli.Text) <> 0 Then
    If retra_llave!TRA_CON4 = 0 Then
      If pub_cadena = "" Then
         pub_cadena = "( {CLIENTES.CLI_RUC_ESPOSO} = '" & LOC_RUC & "')"
      Else
         pub_cadena = pub_cadena + " AND {CLIENTES.CLI_RUC_ESPOSO} = '" & LOC_RUC & "')"
      End If
    Else
      If pub_cadena = "" Then
         pub_cadena = " {CLIENTES.CLI_CODCLIE} = " & Trim(txt_cli.Text)
      Else
         pub_cadena = pub_cadena + " AND " + "{CLIENTES.CLI_CODCLIE} = " & Trim(txt_cli.Text)
      End If
    End If
    If retra_llave!tra_ACT11 = 1 Or retra_llave!TRA_S6 = 1 Then
    
    Else
     If pub_cadena = "" Then
           pub_cadena = pub_cadena + " {CLIENTES.CLI_CP} = '" + loc_cp + "' AND {CLIENTES.CLI_CODCIA} = " + "'" & LK_CODCIA & "' "
     Else
           pub_cadena = pub_cadena + " AND " + " {CLIENTES.CLI_CP} = '" + loc_cp + "' AND {CLIENTES.CLI_CODCIA} = " + "'" & LK_CODCIA & "' "
     End If
    End If
  End If
  If Trim(i_codart2.Text) = "" Then
    WCOD_ORIGINAL = 0
  End If
  If val(WCOD_ORIGINAL) <> 0 Then
   ART_ARTICULO = str(WCOD_ORIGINAL)
  Else
   ART_ARTICULO = ""
  End If
  If (retra_llave!tra_c2 = 1 Or retra_llave!tra_con6 = 1) And WCOD_ORIGINAL <> 0 Then   ' X articulo
   If LK_EMP_PTO = "A" Then
     warma_arti = " {ARTICULO.ARM_CODCIA} = "
   Else
     warma_arti = " {ARTI.ART_CODCIA} = "
   End If
   If pub_cadena = "" Then
       pub_cadena = "{ARTI.ART_KEY} = " & str(WCOD_ORIGINAL)
   Else
       pub_cadena = pub_cadena + " AND " + "{ARTI.ART_KEY} = " & str(WCOD_ORIGINAL)
   End If
   

   If retra_llave!tra_ACT11 = 0 And retra_llave!tra_s12 <> 1 Then
     If pub_cadena = "" Then
          pub_cadena = pub_cadena + warma_arti + "'" & LK_CODCIA & "' "
     Else
          pub_cadena = pub_cadena + " AND " + warma_arti + "'" & LK_CODCIA & "' "
     End If
   End If
  End If
  If retra_llave!tra_con7 = 1 Then ' x FAMI
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
      If retra_llave!tra_ACT11 = 0 And retra_llave!tra_s12 <> 1 Then
       If pub_cadena = "" Then
          pub_cadena = pub_cadena + warma_arti + "'" & LK_CODCIA & "' "
       Else
          pub_cadena = pub_cadena + " AND " + warma_arti & "'" & LK_CODCIA & "' "
       End If
      Else
        xcuenta2 = 1
        CADENITA = ""
        If retra_llave!tra_s12 <> 1 Then
         For fila = 1 To 30
          pu_codcia = Mid(Trim(par_llave!par_art_cias), xcuenta2, 2)
          If Trim(pu_codcia) = "" Then Exit For
          CADENITA = CADENITA + " {ARTI.ART_CODCIA} = '" & pu_codcia & "' OR "
          xcuenta2 = xcuenta2 + 2
         Next fila
        End If
        If CADENITA <> "" Then
          CADENITA = "(" & Mid(CADENITA, 1, Len(CADENITA) - 4) & ")"
          If Trim(pub_cadena) = "" Then
          pub_cadena = pub_cadena + CADENITA
          Else
          pub_cadena = pub_cadena + " AND  " & CADENITA
          End If
        End If
      End If
      wmensa = wmensa + "Fam.: " + wfiltra
  End If
  If retra_llave!TRA_ACT15 = 1 Then ' x LINEAS ART_LINEA
      If LK_EMP_PTO = "A" Then
        warma_arti = " {ARTICULO.ARM_CODCIA} = "
      Else
        warma_arti = " {ARTI.ART_CODCIA} = "
      End If
      GoSub ARMA_LINEAS
      If CADENITA <> "" Then
       If pub_cadena = "" Then
          pub_cadena = pub_cadena + CADENITA
       Else
          pub_cadena = pub_cadena + " AND " + CADENITA
       End If
      End If
      If retra_llave!tra_ACT11 = 0 Then
       If pub_cadena = "" Then
          pub_cadena = pub_cadena + warma_arti + "'" & LK_CODCIA & "' "
       Else
          pub_cadena = pub_cadena + " AND " + warma_arti & "'" & LK_CODCIA & "' "
       End If
      Else
        xcuenta2 = 1
        CADENITA = ""
        For fila = 1 To 30
          pu_codcia = Mid(Trim(par_llave!par_art_cias), xcuenta2, 2)
          If Trim(pu_codcia) = "" Then Exit For
          CADENITA = CADENITA + " {ARTI.ART_CODCIA} = '" & pu_codcia & "' OR "
          xcuenta2 = xcuenta2 + 2
        Next fila
        If CADENITA <> "" Then
          CADENITA = "(" & Mid(CADENITA, 1, Len(CADENITA) - 4) & ")"
          If pub_cadena <> "" Then
           pub_cadena = pub_cadena + " AND  " & CADENITA
          Else
           pub_cadena = pub_cadena + CADENITA
          End If
        End If
      End If
      wmensa = wmensa + "lINEAS: " + wfiltra
   End If
   If retra_llave!TRA_S2 = 1 Then ' x ART_SUBGRU
      If LK_EMP_PTO = "A" Then
        warma_arti = " {ARTICULO.ARM_CODCIA} = "
      Else
        warma_arti = " {ARTI.ART_CODCIA} = "
      End If
      GoSub ARMA_CLASES
      If pub_cadena = "" Then
         pub_cadena = pub_cadena + CADENITA
      Else
        If CADENITA <> "" Then pub_cadena = pub_cadena + " AND " + CADENITA
      End If
      If retra_llave!tra_ACT11 = 0 Then
       If pub_cadena = "" Then
          pub_cadena = pub_cadena + warma_arti + "'" & LK_CODCIA & "' "
       Else
          pub_cadena = pub_cadena + " AND " + warma_arti & "'" & LK_CODCIA & "' "
       End If
      Else
        xcuenta2 = 1
        CADENITA = ""
        For fila = 1 To 30
          pu_codcia = Mid(Trim(par_llave!par_art_cias), xcuenta2, 2)
          If Trim(pu_codcia) = "" Then Exit For
          CADENITA = CADENITA + " {ARTI.ART_CODCIA} = '" & pu_codcia & "' OR "
          xcuenta2 = xcuenta2 + 2
        Next fila
        If CADENITA <> "" Then
          CADENITA = "(" & Mid(CADENITA, 1, Len(CADENITA) - 4) & ")"
          If pub_cadena <> "" Then
             pub_cadena = pub_cadena + " AND  " & CADENITA
          Else
             pub_cadena = pub_cadena + CADENITA
          End If
        End If
      End If
      wmensa = wmensa + "lINEAS: " + wfiltra
      
  ElseIf retra_llave!TRA_CON8 = 1 Then
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
   End If
   If retra_llave!tra_con9 = 1 Then ' x FAMI SUB FAMI
      PUB_TIPREG = 122
      archi = "UPDATE TABLAS SET TAB_CODCLIE = 0 WHERE TAB_CODCIA = '" & LK_CODCIA & "' AND TAB_TIPREG = " & PUB_TIPREG
      CN.Execute archi, rdExecDirect
      If pub_cadena = "" Then
         pub_cadena = "{ARTI.ART_FAMILIA} in [" & str(val(Right(famix.Text, 6))) & "]"
      Else
         pub_cadena = pub_cadena + " AND " + "{ARTI.ART_FAMILIA} in [" & str(val(Right(famix.Text, 6))) & "]"
         PUB_NUMTAB = val(Right(famix.Text, 6))
         archi = "UPDATE TABLAS SET TAB_CODCLIE = 1 WHERE TAB_CODCIA = '" & LK_CODCIA & "' AND TAB_TIPREG = " & PUB_TIPREG & " AND TAB_NUMTAB = " & PUB_NUMTAB
         CN.Execute archi, rdExecDirect
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
     If retra_llave!tra_ACT11 = 0 Then
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
  If retra_llave!TRA_S5 = 1 Then
      warma_arti = " {ARTI.ART_CODCIA} = "
      GoSub ARMA_PLANCHA
      If CADENITA <> "" Then
        If pub_cadena <> "" Then
           pub_cadena = pub_cadena + " AND " + CADENITA
        Else
           pub_cadena = pub_cadena + CADENITA
        End If
      End If
      If retra_llave!tra_ACT11 = 0 Then
       If pub_cadena = "" Then
          pub_cadena = pub_cadena + warma_arti + "'" & LK_CODCIA & "' "
       Else
          pub_cadena = pub_cadena + " AND " + warma_arti & "'" & LK_CODCIA & "' "
       End If
      End If
  End If
  If retra_llave!TRA_ACT13 = 1 Then
      warma_arti = " {ARTI.ART_CODCIA} = "
      GoSub ARMA_MARCA
      If CADENITA <> "" Then
        If pub_cadena <> "" Then
           pub_cadena = pub_cadena + " AND " + CADENITA
        Else
           pub_cadena = pub_cadena + CADENITA
        End If
      End If
      If retra_llave!tra_ACT11 = 0 Then
       If pub_cadena = "" Then
          pub_cadena = pub_cadena + warma_arti + "'" & LK_CODCIA & "' "
       Else
          pub_cadena = pub_cadena + " AND " + warma_arti & "'" & LK_CODCIA & "' "
       End If
      End If
  End If
  
  If retra_llave!TRA_ACT12 = 1 Then
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
      If retra_llave!tra_ACT11 = 0 Then
       If pub_cadena = "" Then
          pub_cadena = pub_cadena + warma_arti + "'" & LK_CODCIA & "' "
       Else
          pub_cadena = pub_cadena + " AND " + warma_arti & "'" & LK_CODCIA & "' "
       End If
      End If
  End If
  
  If retra_llave!TRA_S7 = 1 Then
      If LK_EMP_PTO = "A" Then
        warma_arti = " {ARTICULO.ARM_CODCIA} = "
      Else
        warma_arti = " {ARTI.ART_CODCIA} = "
      End If
      GoSub ARMA_CALIDAD
      If CADENITA <> "" Then
        If pub_cadena <> "" Then
           pub_cadena = pub_cadena + " AND " + CADENITA
        Else
           pub_cadena = pub_cadena + CADENITA
        End If
      End If
      If retra_llave!tra_ACT11 = 0 Then
       If pub_cadena = "" Then
          pub_cadena = pub_cadena + warma_arti + "'" & LK_CODCIA & "' "
       Else
          pub_cadena = pub_cadena + " AND " + warma_arti & "'" & LK_CODCIA & "' "
       End If
      End If
  End If
  
  
  
  If retra_llave!TRA_ACT5 = 1 Then ' x FECHAS X CHEQUES
    If Not SON_FECHAS Then
    GoTo SALE
    End If
    DIA = Day(REP_FECHA1)
    MES = Month(REP_FECHA1)
    ANO = Year(REP_FECHA1)
    DIA1 = Day(REP_FECHA2)
    MES1 = Month(REP_FECHA2)
    ANO1 = Year(REP_FECHA2)
    CADENITA = "{CHEQUES.CHE_FECHA} >= Date ( " & ANO & "," & MES & "," & DIA & ") AND {CHEQUES.CHE_FECHA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
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
  If retra_llave!tra_con1 = 1 Or retra_llave!tra_GRU1 = 1 Then  ' x FECHAS X FACART
    If Not SON_FECHAS Then
        GoTo SALE
    End If
    DIA = Day(REP_FECHA1)
    MES = Month(REP_FECHA1)
    ANO = Year(REP_FECHA1)
    DIA1 = Day(REP_FECHA2)
    MES1 = Month(REP_FECHA2)
    ANO1 = Year(REP_FECHA2)
    If retra_llave!TRA_ACT9 = 1 Then ' x FECHAS X FACART
      pub_mensaje = "Imprimir según Usuario... ?"
      Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
      If Pub_Respuesta = vbYes Then
         CADENITA = "{FACART.FAR_CODUSU}= '" & LK_CODUSU & "' AND {FACART.FAR_FECHA} >= Date ( " & ANO & "," & MES & "," & DIA & ") AND {FACART.FAR_FECHA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
      Else
         If retra_llave!tra_con1 = 1 Then
            CADENITA = "{FACART.FAR_FECHA_COMPRA} >= Date ( " & ANO & "," & MES & "," & DIA & ") AND {FACART.FAR_FECHA_COMPRA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
         Else
            CADENITA = "{FACART.FAR_FECHA} >= Date ( " & ANO & "," & MES & "," & DIA & ") AND {FACART.FAR_FECHA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
         End If
      End If
    Else
      If retra_llave!tra_con1 = 1 Then
          CADENITA = "{FACART.FAR_FECHA_COMPRA} >= Date ( " & ANO & "," & MES & "," & DIA & ") AND {FACART.FAR_FECHA_COMPRA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
      Else
          CADENITA = "{FACART.FAR_FECHA} >= Date ( " & ANO & "," & MES & "," & DIA & ") AND {FACART.FAR_FECHA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
      End If
    End If
    If pub_cadena = "" Then
       pub_cadena = pub_cadena + CADENITA
    Else
      pub_cadena = pub_cadena + " AND " + CADENITA
    End If
    If retra_llave!tra_ACT11 = 0 Then
     If pub_cadena = "" Then
         pub_cadena = pub_cadena + " {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' "
     Else
          pub_cadena = pub_cadena + " AND  {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' "
     End If
    End If
    'Debug.Print pub_cadena
    If LK_FLAG_SOS = "A" Then
       If pub_cadena = "" Then
           pub_cadena = pub_cadena + " {FACART.FAR_FLAG_SO} = 'A' "
       Else
           pub_cadena = pub_cadena + " AND {FACART.FAR_FLAG_SO} = 'A' "
       End If
    End If
  End If
  If retra_llave!tra_s12 = 1 Or retra_llave!TRA_S9 = 1 Or retra_llave!tra_ACT11 = 1 Or retra_llave!TRA_S6 = 1 Then
        xcuenta2 = 1
        CADENITA = ""
        cade2 = ""
        wsalmacenes = ""
        wfiltra = ""
        wran1 = ""
        wran2 = ""
        For fila = 0 To LISCIA.ListCount - 1
          LISCIA.ListIndex = fila
          pu_codcia = Left(LISCIA.Text, 2)
          wran1 = wran1 + "FAR_CODCIA = '" & pu_codcia & "'"
          If retra_llave!TRA_S6 = 1 Then
            cade2 = cade2 + " {CARTERA.CAR_CODCIA} = '" & pu_codcia & "' OR "
          ElseIf retra_llave!TRA_S9 = 1 Then
            cade2 = cade2 + " {ALLOG.ALL_CODCIA} = '" & pu_codcia & "' OR "
          ElseIf retra_llave!tra_s12 = 1 Then
            cade2 = cade2 + " {ARTICULO.ARM_CODCIA} = '" & pu_codcia & "' OR "
          Else
            cade2 = cade2 + " {FACART.FAR_CODCIA} = '" & pu_codcia & "' OR "
          End If
          ''wfiltra = wfiltra + pu_codcia & " - "
          PSPAR_MULTI(0) = pu_codcia
          par_multi.Requery
          wfiltra = wfiltra + Trim(par_multi!par_nombre_corto) & " - "
          If LISCIA.Selected(fila) Then
             If retra_llave!TRA_S6 = 1 Then
               CADENITA = CADENITA + " {CARTERA.CAR_CODCIA} = '" & pu_codcia & "' OR "
             ElseIf retra_llave!TRA_S9 = 1 Then
               CADENITA = CADENITA + " {ALLOG.ALL_CODCIA} = '" & pu_codcia & "' OR "
             ElseIf retra_llave!tra_s12 = 1 Then
               CADENITA = CADENITA + " {ARTICULO.ARM_CODCIA} = '" & pu_codcia & "' OR "
             Else
               CADENITA = CADENITA + " {FACART.FAR_CODCIA} = '" & pu_codcia & "' OR "
             End If
             wran2 = wran2 + "FAR_CODCIA = '" & pu_codcia & "'"
            '' PSPAR_MULTI(0) = pu_codcia
            '' par_multi.Requery
             wsalmacenes = wsalmacenes + Trim(par_multi!par_nombre_corto) & " - "
          End If
        Next fila
        If cade2 <> "" Then
          cade2 = "(" & Mid(cade2, 1, Len(cade2) - 4) & ")"
          If Trim(wsalmacenes) <> "" Then
            wfiltra = Trim(GEN!GEN_NOMBRE) & " " & Mid(wsalmacenes, 1, Len(wsalmacenes) - 3)
          Else
            wfiltra = Trim(GEN!GEN_NOMBRE) & " " & Mid(wfiltra, 1, Len(wfiltra) - 3)
          End If
          wran1 = "(" & Mid(wran1, 1, Len(wran1) - 3) & ")"
        End If
        If LISCIA.ListCount = 0 Then
            pu_codcia = LK_CODCIA
            If retra_llave!TRA_S6 = 1 Then
                CADENITA = CADENITA + " {CARTERA.CAR_CODCIA} = '" & pu_codcia & "' OR "
            ElseIf retra_llave!TRA_S9 = 1 Then
                CADENITA = CADENITA + " {ALLOG.ALL_CODCIA} = '" & pu_codcia & "' OR "
            ElseIf retra_llave!tra_s12 = 1 Then
                CADENITA = CADENITA + " {ARTICULO.ARM_CODCIA} = '" & pu_codcia & "' OR "
            Else
                CADENITA = CADENITA + " {FACART.FAR_CODCIA} = '" & pu_codcia & "' OR "
            End If
            wran2 = wran2 + "FAR_CODCIA = '" & pu_codcia & "'"
        End If
        
        If CADENITA <> "" Then
          CADENITA = "(" & Mid(CADENITA, 1, Len(CADENITA) - 4) & ")"
          wran2 = "(" & Mid(wran2, 1, Len(wran2) - 3) & ")"
          If pub_cadena <> "" Then
           pub_cadena = pub_cadena + " AND  " & CADENITA
          Else
           pub_cadena = pub_cadena + CADENITA
          End If
         wsalmacenes = wfiltra
          ''wsalmacenes = Trim(par_llave!par_nombre_corto) & " Almacen : " & Mid(wsalmacenes, 1, Len(wsalmacenes) - 3)
        Else
          CADENITA = cade2
          If CADENITA <> "" Then
           If pub_cadena <> "" Then
             pub_cadena = pub_cadena + " AND  " & CADENITA
           Else
             pub_cadena = pub_cadena + CADENITA
           End If
          End If
          wsalmacenes = wfiltra
        End If
    End If
  If retra_llave!TRA_CON14 = 1 Then ' x FECHAS X ALLOG
    If Not SON_FECHAS Then
     Exit Sub
    End If
    DIA = Day(REP_FECHA1)
    MES = Month(REP_FECHA1)
    ANO = Year(REP_FECHA1)
    DIA1 = Day(REP_FECHA2)
    MES1 = Month(REP_FECHA2)
    ANO1 = Year(REP_FECHA2)
    CADENITA = "{ALLOG.ALL_FECHA_SUNAT} >= Date ( " & ANO & "," & MES & "," & DIA & ") AND {ALLOG.ALL_FECHA_SUNAT} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
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
  If retra_llave!TRA_S8 = 1 Then
    If Not SON_FECHAS Then
     Exit Sub
    End If
    DIA = Day(REP_FECHA1)
    MES = Month(REP_FECHA1)
    ANO = Year(REP_FECHA1)
    DIA1 = Day(REP_FECHA2)
    MES1 = Month(REP_FECHA2)
    ANO1 = Year(REP_FECHA2)
    CADENITA = "{ALLOG.ALL_FECHA_PRO} >= Date ( " & ANO & "," & MES & "," & DIA & ") AND {ALLOG.ALL_FECHA_PRO} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
    If pub_cadena = "" Then
       pub_cadena = pub_cadena + CADENITA
    Else
      pub_cadena = pub_cadena + " AND " + CADENITA
    End If
    If retra_llave!TRA_S9 <> 1 Then
     If pub_cadena = "" Then
          pub_cadena = pub_cadena + " {ALLOG.ALL_CODCIA} = '" & LK_CODCIA & "' "
     Else
          pub_cadena = pub_cadena + " AND  {ALLOG.ALL_CODCIA} = '" & LK_CODCIA & "' "
     End If
    End If
  End If
  If retra_llave!TRA_S14 = 1 Then
    CADENITA = "{ALLOG.ALL_MONEDA_CAJA} = '" & Left(cmdmoneda.Text, 1) & "'"
    If pub_cadena = "" Then
       pub_cadena = pub_cadena + CADENITA
    Else
      pub_cadena = pub_cadena + " AND " + CADENITA
    End If
  End If
  If retra_llave!TRA_S15 = 1 Then
    If retra_llave!TRA_KEY <> 9306 Then ' Modificado 20042004
    CADENITA = "{CARTERA.CAR_MONEDA} = '" & Left(cmdmoneda.Text, 1) & "'"
    If pub_cadena = "" Then
       pub_cadena = pub_cadena + CADENITA
    Else
      pub_cadena = pub_cadena + " AND " + CADENITA
    End If
    End If
  End If
  If retra_llave!TRA_C1 = 1 Then
    CADENITA = "{FACART.FAR_MONEDA} = '" & Left(cmdmoneda.Text, 1) & "'"
    If pub_cadena = "" Then
       pub_cadena = pub_cadena + CADENITA
    Else
      pub_cadena = pub_cadena + " AND " + CADENITA
    End If
  End If
  If retra_llave!TRA_S11 = 1 Then
    If Not SON_FECHAS Then
     Exit Sub
    End If
    DIA = Day(REP_FECHA1)
    MES = Month(REP_FECHA1)
    ANO = Year(REP_FECHA1)
    DIA1 = Day(REP_FECHA2)
    MES1 = Month(REP_FECHA2)
    ANO1 = Year(REP_FECHA2)
    CADENITA = "{ALLOG.ALL_FECHA_CAN} >= Date ( " & ANO & "," & MES & "," & DIA & ") AND {ALLOG.ALL_FECHA_CAN} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
    If pub_cadena = "" Then
       pub_cadena = pub_cadena + CADENITA
    Else
      pub_cadena = pub_cadena + " AND " + CADENITA
    End If
    If retra_llave!TRA_S9 <> 1 Then
     If pub_cadena = "" Then
          pub_cadena = pub_cadena + " {ALLOG.ALL_CODCIA} = '" & LK_CODCIA & "' "
     Else
          pub_cadena = pub_cadena + " AND  {ALLOG.ALL_CODCIA} = '" & LK_CODCIA & "' "
     End If
    End If
  End If
  If retra_llave!tra_con10 = 1 Then ' x FECHA DE CARTERA VCTO
    If Not SON_FECHAS Then
      GoTo SALE
     Exit Sub
    End If
    DIA = Day(REP_FECHA1)
    MES = Month(REP_FECHA1)
    ANO = Year(REP_FECHA1)
    DIA1 = Day(REP_FECHA2)
    MES1 = Month(REP_FECHA2)
    ANO1 = Year(REP_FECHA2)
    CADENITA = "{CARTERA.CAR_FECHA_VCTO} >= Date ( " & ANO & "," & MES & "," & DIA & ") AND {CARTERA.CAR_FECHA_VCTO} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
    If pub_cadena = "" Then
       pub_cadena = pub_cadena + CADENITA
    Else
      pub_cadena = pub_cadena + " AND " + CADENITA
    End If
    If retra_llave!TRA_S6 = 0 Then ' x FECHA DE CARTERA
     If pub_cadena = "" Then
          pub_cadena = pub_cadena + " {CARTERA.CAR_CODCIA} = '" & LK_CODCIA & "' "
     Else
          pub_cadena = pub_cadena + " AND  {CARTERA.CAR_CODCIA} = '" & LK_CODCIA & "' "
     End If
    End If
    If LK_FLAG_SOS = "A" Then
       If pub_cadena = "" Then
           pub_cadena = pub_cadena + " {CARTERA.CAR_FLAG_SO} = 'A' "
       Else
           pub_cadena = pub_cadena + " AND {CARTERA.CAR_FLAG_SO} = 'A' "
       End If
    End If
  End If
  If retra_llave!tra_act8 = 1 Then ' x FECHA DE CARTERA
    If Not SON_FECHAS Then
      GoTo SALE
     Exit Sub
    End If
    DIA = Day(REP_FECHA1)
    MES = Month(REP_FECHA1)
    ANO = Year(REP_FECHA1)
    DIA1 = Day(REP_FECHA2)
    MES1 = Month(REP_FECHA2)
    ANO1 = Year(REP_FECHA2)
    If Wfile = "PLANINI.RPT" Then
       CADENITA = "{CARTERA.CAR_FECHA_INGR} >= Date ( " & ANO & "," & MES & "," & DIA & ") AND {CARTERA.CAR_FECHA_INGR} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
    Else
       CADENITA = "{CARTERA.CAR_FECHA_SUNAT} >= Date ( " & ANO & "," & MES & "," & DIA & ") AND {CARTERA.CAR_FECHA_SUNAT} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
    End If
    If pub_cadena = "" Then
       pub_cadena = pub_cadena + CADENITA
    Else
      pub_cadena = pub_cadena + " AND " + CADENITA
    End If
    If retra_llave!TRA_S6 = 0 Then ' x FECHA DE CARTERA
     If pub_cadena = "" Then
          pub_cadena = pub_cadena + " {CARTERA.CAR_CODCIA} = '" & LK_CODCIA & "' "
     Else
          pub_cadena = pub_cadena + " AND  {CARTERA.CAR_CODCIA} = '" & LK_CODCIA & "' "
     End If
    End If
    If LK_FLAG_SOS = "A" Then
       If pub_cadena = "" Then
           pub_cadena = pub_cadena + " {CARTERA.CAR_FLAG_SO} = 'A' "
       Else
           pub_cadena = pub_cadena + " AND {CARTERA.CAR_FLAG_SO} = 'A' "
       End If
    End If
  End If
  If retra_llave!tra_con12 = 1 Then ' x FECHA DE CARACU
    If Not SON_FECHAS Then
     Exit Sub
    End If
    DIA = Day(REP_FECHA1)
    MES = Month(REP_FECHA1)
    ANO = Year(REP_FECHA1)
    DIA1 = Day(REP_FECHA2)
    MES1 = Month(REP_FECHA2)
    ANO1 = Year(REP_FECHA2)
    CADENITA = "{CARACU.CAA_FECHA_COBRO} >= Date ( " & ANO & "," & MES & "," & DIA & ") AND {CARACU.CAA_FECHA_COBRO} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
    If pub_cadena = "" Then
       pub_cadena = pub_cadena + CADENITA
    Else
      pub_cadena = pub_cadena + " AND " + CADENITA
    End If
    If retra_llave!TRA_S6 = 0 Then
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
  If retra_llave!TRA_CON11 = 1 Then
     CADENITA = " {CLIENTES.CLI_CODCIA} = '" & LK_CODCIA & "' AND {CLIENTES.CLI_CP} = '" & Left(cmbclipro.Text, 1) & "'"
     If pub_cadena = "" Then
         pub_cadena = pub_cadena + CADENITA
     Else
        If CADENITA <> "" Then
         pub_cadena = pub_cadena + " AND " + CADENITA
        End If
     End If
  End If
  If retra_llave!TRA_CON3 = 1 Then ' x VENDEDOR
      GoSub ARMA_VEND:
      If pub_cadena = "" Then
         pub_cadena = pub_cadena + CADENITA
      Else
        If CADENITA <> "" Then
         pub_cadena = pub_cadena + " AND " + CADENITA
        End If
      End If
      If retra_llave!tra_ACT11 <> 1 Then
        If pub_cadena = "" Then
          pub_cadena = pub_cadena + " {VEMAEST.VEM_CODCIA} = '" & LK_CODCIA & "' "
        Else
          pub_cadena = pub_cadena + " AND  {VEMAEST.VEM_CODCIA} = '" & LK_CODCIA & "' "
        End If
      End If
      wmensa = wmensa + "Ven.: " + wfiltra
  End If
  If retra_llave!TRA_ACT4 = 1 Then ' x VENDEDOR
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
  If retra_llave!TRA_S13 = 1 Then ' x DIVISION
      GoSub ARMA_DIVISION
      If pub_cadena = "" Then
         pub_cadena = pub_cadena + CADENITA
      Else
        If CADENITA <> "" Then
         pub_cadena = pub_cadena + " AND " + CADENITA
        End If
      End If
  End If

  If retra_llave!TRA_ACT6 = 1 Then
    If pub_cadena = "" Then
       pub_cadena = "{ARTI.ART_MONEDA} = '" & Left(cmdmoneda.Text, 1) & "'"
    Else
        pub_cadena = pub_cadena + " AND " + "{ARTI.ART_MONEDA} = '" & Left(cmdmoneda.Text, 1) & "'"
    End If
  End If
  If retra_llave!TRA_ACT7 = 1 Then
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
  If retra_llave!TRA_ACT14 = 1 Then
    WFECHA_DIA = Format(LK_FECHA_DIA, "dd/mm/") & Format((val(Format(LK_FECHA_DIA, "yyyy")) - 6), "####")
  Else
    WFECHA_DIA = Format(LK_FECHA_DIA, "dd/mm/yyyy")
  End If
  If retra_llave!tra_ACT1 = 1 Then
    wf1 = "CIA=  '" & Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))) & "'"
    If retra_llave!tra_s12 = 1 Or retra_llave!tra_ACT11 = 1 Or retra_llave!TRA_S6 = 1 Then
     wf1 = "CIA=  '" & wsalmacenes & "'"
    End If
  End If
  If retra_llave!tra_ACT2 = 1 Then
    wf2 = "DIA=  '" & WFECHA_DIA & "'"
  End If
  If retra_llave!tra_ACT10 = 1 Then
    wf3 = "FECHAS=  ' DEL " & REP_FECHA1 & " AL " & REP_FECHA2 & "'"
  End If
  If retra_llave!TRA_CON15 = 1 Then ' X para fecha de rango en columnas 12 maximos
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
  Reportes.Formulas(20) = ""
  Reportes.Formulas(21) = ""
  If LK_EMP = "HER" And (Wfile = "CLIVENTA.RPT" Or Wfile = "CLILINVEN.RPT") Then
    Reportes.Formulas(20) = "RANGO1=" & str(val(op1.Text))
    Reportes.Formulas(21) = "RANGO2=" & str(val(op2.Text))
  End If
  Reportes.Formulas(50) = ""
  If retra_llave!TRA_ACT3 = 1 Then
      DIA = Day(LK_FECHA_DIA)
      MES = Month(LK_FECHA_DIA)
      ANO = Year(LK_FECHA_DIA)
      Reportes.Formulas(50) = "FECHADIA= Date ( " & ANO & "," & MES & "," & DIA & ")"
  End If
  Reportes.Formulas(51) = ""
  If retra_llave!TRA_S10 = 1 Then
      Reportes.Formulas(51) = "FLAG= " & str(cheflag.Value)
  End If
  Reportes.Formulas(52) = ""
  If retra_llave!TRA_GRU2 = 1 Then
      Reportes.Formulas(52) = "FLAG2= " & str(CHEFLAG2.Value)
  End If
  'pub_cadena = "{CCMAEST.CCM_CODBAN} = 104 AND  ( {ALLOG.ALL_CODCIA} = '01' OR  {ALLOG.ALL_CODCIA} = '02') AND {ALLOG.ALL_FECHA_PRO} >= Date ( 2001,6,1) AND {ALLOG.ALL_FECHA_PRO} <= Date ( 2001,6,30)"
  If Wfile = "CONSO1.RPT" Then
    CADENITA = ""
    wfiltra = ""
    Modo1 = "{AUTORIZACION.AUT_CODCLIE} in ["
    For fila = 0 To multiven.ListCount - 1
          multiven.ListIndex = fila
          If multiven.Selected(fila) Then
            wfiltra = wfiltra + Left(multiven.Text, 3) + ","
            Modo1 = Modo1 + str(val(Left(multiven.Text, 3))) + ","
          End If
    Next fila
    If wfiltra <> "" Then
          CADENITA = Left(Modo1, Len(Modo1) - 1) & "] "
          wfiltra = Left(wfiltra, Len(wfiltra) - 1)
    Else
          CADENITA = ""
          wfiltra = "(*)"
     End If
     pub_cadena = "{AUTORIZACION.AUT_FECHA} >= Date ( " & ANO & "," & MES & "," & DIA & ") AND {AUTORIZACION.AUT_FECHA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ") "
     If CADENITA <> "" Then
     pub_cadena = pub_cadena & " AND " & CADENITA
     End If
  End If
  If retra_llave("tra_key") = 9304 Or retra_llave("tra_key") = 9311 Then 'agregado por mic
    If val(Right(cboDias.Text, 6)) <> 0 And val(Right(cboDias.Text, 6)) <> 9 And val(Right(cboDias.Text, 6)) <> 8 Then
        Reportes.Formulas(4) = "DIAVISITA= 'Dia Visita : " & Trim(Left(cboDias.Text, 40)) & "'"
        pub_cadena = pub_cadena & " AND {clientes.cli_prenda} = '" & val(Right(cboDias.Text, 6)) & "'"
    End If
    ANO = Year(WFECHA_DIA)
    MES = Month(WFECHA_DIA)
    DIA = Day(WFECHA_DIA)
    Reportes.Formulas(5) = "FECACTUAL = Date (" & ANO & "," & MES & "," & DIA & ")"
  ElseIf retra_llave("tra_key") = 9118 Then 'agregado por mic
    If val(Right(cboDias.Text, 6)) <> 0 And val(Right(cboDias.Text, 6)) <> 9 And val(Right(cboDias.Text, 6)) <> 8 Then
        Reportes.Formulas(4) = "DIAVISITA= 'Dia Visita : " & Trim(Left(cboDias.Text, 40)) & "'"
        pub_cadena = pub_cadena & " AND {clientes.cli_prenda} = '" & val(Right(cboDias.Text, 6)) & "'"
    End If
  ElseIf retra_llave("tra_key") = 9302 Or retra_llave("tra_key") = 8132 Or retra_llave("tra_key") = 8133 Or retra_llave("tra_key") = 8522 Or retra_llave("tra_key") = 8531 Then  'agregado por gts para filtrar usuarios en reporte de caja
     If Trim(cboUsers.Text) <> "" Then
        pub_cadena = pub_cadena & " AND {allog.ALL_CODUSU} = '" & Trim(cboUsers.Text) & "'"
     End If
  ElseIf retra_llave("tra_key") = 9303 Then  'agregado por gts para filtrar usuarios en reporte de caja
     If Trim(cboUsers.Text) <> "" Then
        pub_cadena = pub_cadena & " AND {allog.ALL_CODUSU} = '" & Trim(cboUsers.Text) & "'"
     End If
  ElseIf retra_llave("tra_key") = 8121 Or retra_llave("tra_key") = 8716 Or retra_llave("tra_key") = 8116 Or retra_llave("tra_key") = 8134 Or retra_llave("tra_key") = 8510 Or retra_llave("tra_key") = 8716 Then     'agregado por gts para filtrar usuarios en reporte de ventas de articulos
     If Trim(cboUsers.Text) <> "" Then
        pub_cadena = pub_cadena & " AND {facart.FAR_CODUSU} = '" & Trim(cboUsers.Text) & "'"
     End If
  ElseIf retra_llave("tra_key") = 8125 Then 'agregado por micFAR_PORDESCTOS
        pub_cadena = pub_cadena & " AND {vw_clivtasublinea.SubTotal} > " & op1.Text & " AND {vw_clivtasublinea.SubTotal} < " & op2.Text
        pub_cadena = Replace(pub_cadena, "ARTI", "vw_clivtasublinea")
        pub_cadena = Replace(pub_cadena, "FACART", "vw_clivtasublinea")
  ElseIf retra_llave("tra_key") = 8126 Then 'agregado por mic
        pub_cadena = pub_cadena & " AND {FACART.FAR_PORDESCTOS} = '" & val(txtporcentaje) & "'"
        Reportes.Formulas(3) = "PORCENTAJE='" & (txtporcentaje) & "'"
        Reportes.Formulas(4) = "TITULO='" & Trim(retra_llave("TRA_DESCRIPCION")) & "'"
  End If
  Reportes.SelectionFormula = pub_cadena
  If x = 1 Then
     Exit Sub
  End If
  'pub_cadena = "{CCMAEST.CCM_CODBAN} = 928 AND  ( {ALLOG.ALL_CODCIA} = '01' OR  {ALLOG.ALL_CODCIA} = '02') AND ({ALLOG.ALL_FECHA_CAN} >= Date ( 2001,6,1) AND {ALLOG.ALL_FECHA_CAN} <= Date ( 2001,6,30))"
   '  Debug.Print pub_cadena
  ART_CLASES = ""
  Reportes.Action = 1
  
  ProgBar.Value = 10
  Screen.MousePointer = 0
  ProgBar.Visible = False
  lblproceso.Visible = False
  Pantalla.Enabled = True
  cmdCerrar.Enabled = True

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
    Modo1 = Modo1 + str(val(Right(zonas.Text, 6))) + ","
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
Modo1 = ""
CADENITA = ""
wfiltra = ""
If Nulo_Valor0(retra_llave!TRA_CON13) = 1 Then
  Modo1 = "{FAMILIA.TAB_NUMTAB} in ["
Else
  Modo1 = "{ARTI.ART_FAMILIA} in ["
End If
For fila = 0 To fami.ListCount - 1
  fami.ListIndex = fila
  If fami.Selected(fila) Then
    wfiltra = wfiltra + Left(fami.Text, 8) + ","
    Modo1 = Modo1 + str(val(Right(fami.Text, 6))) + ","
  End If
Next fila
If wfiltra <> "" Then
  CADENITA = Left(Modo1, Len(Modo1) - 1) & "] "
  wfiltra = Left(wfiltra, Len(wfiltra) - 1)
Else
  CADENITA = ""
  wfiltra = "(*)"
End If

If Nulo_Valor0(retra_llave!TRA_CON13) = 1 Then
  If CADENITA <> "" Then
     CADENITA = CADENITA + " AND {FAMILIA.TAB_TIPREG} = 122 AND {FAMILIA.TAB_CODCIA} = '" & wcodcia & "' "
  Else
     CADENITA = "{FAMILIA.TAB_TIPREG} = 122 AND {FAMILIA.TAB_CODCIA} = '" & wcodcia & "' "
  End If
End If

Return

ARMA_LINEAS:
PUB_TIPREG = 131
archi = "UPDATE TABLAS SET TAB_CODCLIE = 0 WHERE TAB_CODCIA = '" & LK_CODCIA & "' AND TAB_TIPREG = " & PUB_TIPREG
CN.Execute archi, rdExecDirect

Modo1 = ""
CADENITA = ""
wfiltra = ""
ART_CLASES = ""
If val(Right(LINEAS.Text, 8)) <> 0 Then
   CADENITA = "{ARTI.ART_LINEA} in [" & Trim(Right(LINEAS.Text, 8)) & "]"
   PUB_NUMTAB = val(Right(LINEAS.Text, 8))
   archi = "UPDATE TABLAS SET TAB_CODCLIE = 1 WHERE TAB_CODCIA = '" & LK_CODCIA & "' AND TAB_TIPREG = " & PUB_TIPREG & " AND TAB_NUMTAB = " & PUB_NUMTAB
   CN.Execute archi, rdExecDirect
Else
   CADENITA = ""
End If
 
'End If

'For fila = 0 To LINEAS.ListCount - 1
'  LINEAS.ListIndex = fila
  'If LINEAS.Selected(fila) Then
  '  wfiltra = wfiltra + Left(LINEAS.Text, 8) + ","
  '  Modo1 = Modo1 + str(Val(Right(LINEAS.Text, 6))) + ","
  '  ART_CLASES = ART_CLASES + str(Val(Right(LINEAS.Text, 6))) + ","
 ' End If
'Next fila
'If wfiltra <> "" Then
'  CADENITA = Left(Modo1, Len(Modo1) - 1) & "] "
'  wfiltra = Left(wfiltra, Len(wfiltra) - 1)
'  ART_CLASES = Left(ART_CLASES, Len(ART_CLASES) - 1) & ") "
'Else
'  CADENITA = ""
'  ART_CLASES = ""
'  wfiltra = "(*)"
'End If

'If Nulo_Valor0(retra_llave!TRA_CON13) = 1 Then
'  If CADENITA <> "" Then
'     CADENITA = CADENITA + " AND {FAMILIA.TAB_TIPREG} = 129 AND {FAMILIA.TAB_CODCIA} = '" & wcodcia & "' "
'  Else
'     CADENITA = "{FAMILIA.TAB_TIPREG} = 129 AND {FAMILIA.TAB_CODCIA} = '" & wcodcia & "' "
'  End If
'End If

Return


ARMA_CLASES:
PUB_TIPREG = 129
archi = "UPDATE TABLAS SET TAB_CODCLIE = 0 WHERE TAB_CODCIA = '" & LK_CODCIA & "' AND TAB_TIPREG = " & PUB_TIPREG
CN.Execute archi, rdExecDirect

CADENITA = ""
If val(Right(art_subgru.Text, 8)) <> 0 Then
   CADENITA = "{ARTI.ART_SUBGRU} in [" & Trim(Right(art_subgru.Text, 8)) & "]"
   PUB_NUMTAB = val(Right(art_subgru.Text, 8))
   archi = "UPDATE TABLAS SET TAB_CODCLIE = 1 WHERE TAB_CODCIA = '" & LK_CODCIA & "' AND TAB_TIPREG = " & PUB_TIPREG & " AND TAB_NUMTAB = " & PUB_NUMTAB
   CN.Execute archi, rdExecDirect

Else
   CADENITA = ""
End If

Return



ARMA_NUMERO:
PUB_TIPREG = 130
archi = "UPDATE TABLAS SET TAB_CODCLIE = 0 WHERE TAB_CODCIA = '" & LK_CODCIA & "' AND TAB_TIPREG = " & PUB_TIPREG
CN.Execute archi, rdExecDirect

CADENITA = ""
Modo1 = "{ARTI.ART_NUMERO} in ["
If val(Right(art_numero.Text, 8)) <> 0 Then
   CADENITA = "{ARTI.ART_NUMERO} in [" & Trim(Right(art_numero.Text, 8)) & "]"
   PUB_NUMTAB = val(Right(art_numero.Text, 8))
   archi = "UPDATE TABLAS SET TAB_CODCLIE = 1 WHERE TAB_CODCIA = '" & LK_CODCIA & "' AND TAB_TIPREG = " & PUB_TIPREG & " AND TAB_NUMTAB = " & PUB_NUMTAB
   CN.Execute archi, rdExecDirect
Else
   CADENITA = ""
End If
Return

ARMA_MARCA:
PUB_TIPREG = 132
archi = "UPDATE TABLAS SET TAB_CODCLIE = 0 WHERE TAB_CODCIA = '" & LK_CODCIA & "' AND TAB_TIPREG = " & PUB_TIPREG
CN.Execute archi, rdExecDirect

CADENITA = ""
Modo1 = "{ARTI.ART_MARCA} in ["
If val(Right(art_marca.Text, 8)) <> 0 Then
   CADENITA = "{ARTI.ART_MARCA} in [" & Trim(Right(art_marca.Text, 8)) & "]"
   PUB_NUMTAB = val(Right(art_marca.Text, 8))
   archi = "UPDATE TABLAS SET TAB_CODCLIE = 1 WHERE TAB_CODCIA = '" & LK_CODCIA & "' AND TAB_TIPREG = " & PUB_TIPREG & " AND TAB_NUMTAB = " & PUB_NUMTAB
   CN.Execute archi, rdExecDirect
Else
   CADENITA = ""
End If
Return

ARMA_PLANCHA:
CADENITA = ""
Modo1 = "{ARTI.ART_PLANCHA} in ["
If val(Right(art_plancha.Text, 8)) <> 0 Then
   CADENITA = "{ARTI.ART_PLANCHA} in [" & Trim(Right(art_plancha.Text, 8)) & "]"
Else
   CADENITA = ""
End If
Return

ARMA_CALIDAD:
CADENITA = ""
wfiltra = ""
Modo1 = "{ARTI.ART_CALIDAD} in ["
For fila = 0 To listacal.ListCount - 1
  listacal.ListIndex = fila
  If listacal.Selected(fila) Then
    wfiltra = wfiltra + Left(listacal.Text, 8) + ","
    Modo1 = Modo1 + str(val(Right(listacal.Text, 6))) + ","
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
PUB_TIPREG = 123
archi = "UPDATE TABLAS SET TAB_CODCLIE = 0 WHERE TAB_CODCIA = '" & LK_CODCIA & "' AND TAB_TIPREG = " & PUB_TIPREG
CN.Execute archi, rdExecDirect
CADENITA = ""
wfiltra = ""
CADENITA = ""
If val(Right(subfami.Text, 8)) <> 0 Then
   CADENITA = "{ARTI.ART_SUBFAM} in [" & Trim(Right(subfami.Text, 8)) & "]"
   PUB_NUMTAB = val(Right(subfami.Text, 8))
   archi = "UPDATE TABLAS SET TAB_CODCLIE = 1 WHERE TAB_CODCIA = '" & LK_CODCIA & "' AND TAB_TIPREG = " & PUB_TIPREG & " AND TAB_NUMTAB = " & PUB_NUMTAB
   CN.Execute archi, rdExecDirect

Else
   CADENITA = ""
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
    Modo1 = Modo1 + str(val(Left(multiven.Text, 3))) + ","
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

ARMA_DIVISION:
CADENITA = ""
wfiltra = ""
Modo1 = "{CLIENTES.CLI_DIVISION} in ["
For fila = 0 To lsttipo.ListCount - 1
  lsttipo.ListIndex = fila
  If lsttipo.Selected(fila) Then
    wfiltra = wfiltra + Trim(Right(lsttipo.Text, 6)) + ","
    Modo1 = Modo1 + Trim(Right(lsttipo.Text, 6)) + " ,"
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
ANO = Year(REP_FECHA1)
ANO1 = Year(REP_FECHA2)
If ANO = ANO1 Then
  Reportes.Formulas(11) = "ANO = '" & ANO & "'"
Else
  Reportes.Formulas(11) = "ANO = '" & ANO & " - " & ANO1 & "'"
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
    Reportes.Formulas(12 + xcuenta) = "A" & i & "=" & ANO
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
  Debug.Print pub_cadena
' Resume Next
 Pantalla.Enabled = True
 cmdCerrar.Enabled = True

End Sub


Private Sub subfami_Click()
WFAMI2 = val(Trim(Right(subfami.Text, 6)))
PUB_TIPREG = 129
LLENADO_SUBFAM art_subgru, WFAMI2
art_numero.Clear
LINEAS.Clear
art_marca.Clear


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
Dim VAR As String
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
If KeyAscii = 13 And Left(txt_cli.Text, 1) = "+" Then GoTo buscar
On Error GoTo ERROR_CODIGO
pu_codclie = val(txt_cli.Text)
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
     lblCliente.Caption = Trim(cli_llave!CLI_NOMBRE)
     LOC_RUC = Trim(cli_llave!cli_ruc_esposo)
   End If
   If Pantalla.Visible And Pantalla.Enabled Then
     Pantalla.SetFocus
   End If
Else
   If loc_key > ListView2.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   valor = UCase(ListView2.ListItems.Item(loc_key).Text)
   If Trim(UCase(txt_cli.Text)) = Left(valor, Len(Trim(txt_cli.Text))) Then
   Else
      Exit Sub
   End If
   txt_cli.Text = Trim(ListView2.ListItems.Item(loc_key).SubItems(1))
   pu_codclie = val(txt_cli.Text)
   SQ_OPER = 1
   pu_cp = loc_cp
   pu_codcia = LK_CODCIA
   LEER_CLI_LLAVE
   If Not cli_llave.EOF Then
    lblCliente.Caption = Trim(ListView2.ListItems.Item(loc_key).Text)
    LOC_RUC = Trim(cli_llave!cli_ruc_esposo)
   End If
   
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
Exit Sub

buscar:
VAR = Mid(txt_cli.Text, 2, Len(txt_cli.Text))
numarchi = alta_vista_nombre(ListView2, VAR, loc_cp)
If numarchi = 0 Then
  ListView2.Visible = False
  MsgBox "Alta Vista: No Existe .. Esta descripcion..", 48, Pub_Titulo
Else
  ListView2.Visible = True
  txt_cli.SetFocus
End If
loc_key = 1


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
    'archi = "SELECT CLI_CODCLIE, CLI_CODCIA, CLI_CP, CLI_NOMBRE,CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM  FROM CLIENTES WHERE  CLI_CP = '" & loc_cp & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & txt_cli.Text & "' AND  '" & VAR & "' ORDER BY CLI_NOMBRE"
    archi = "SELECT CLI_CODCLIE , CLI_CODCIA, CLI_CP, CLI_NOMBRE, CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM, TAB_NOMLARGO  FROM CLIENTES,TABLAS WHERE (TAB_CODCIA = '00') AND (TAB_TIPREG = 35) AND (TAB_NUMTAB = CLI_ZONA_NEW) AND CLI_CP = '" & loc_cp & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & txt_cli.Text & "' AND  '" & VAR & "' ORDER BY CLI_NOMBRE"
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

Private Sub Txt_key_Change()
If Trim(txt_key.Text) = "" Then lblbanco.Caption = ""
End Sub

Private Sub txtCampo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Azul2 txtCampo2, txtCampo2
End If
End Sub
Public Sub KARDEX_CLASES()
'On Error GoTo FINTODO
Dim WCAT As Currency
Dim q_stock_val As Currency
Dim TOTAL_CLASE As Currency
Dim TOTAL_CLASE_VAL As Currency
Dim WCONCIA As Integer
Dim Wfecha_resulta As Date
Dim CHE_KARDEX As Currency
Dim WCOSPRO_SUP As Currency
Dim WTC As Currency
Dim wtotal As Currency
Dim WCIA1 As String * 2
Dim WCIA2 As String * 2
Dim WCIA3 As String * 2
Dim WCIA4 As String * 2
Dim WSCODART As Currency
Dim flag_xx As Integer
Dim ww_concepto As String
Dim ww_tipo As String
Dim ww_codcia As String * 2
Dim WS_PRECIO As Currency
Dim WW_LINEA, i
Dim ws_clave As String
Dim FF1 As Long
Dim LETRAS(24) As String * 1
Dim WSFECHA As Date
Dim acu_val_ingresos As Currency
Dim acu_val_salidas As Currency
Dim xx_ingreso As Currency
Dim xx_salida As Currency
Dim ww_ingreso As Currency
Dim ww_salida As Currency
Dim acu_cant_dia As Currency
Dim acu_saldo As Currency
Dim acu_stock As Currency
Dim wsfile As String
Dim walterno As String
Dim wdnombre As String
Dim WD_COSPRO As Currency
Dim q_sum_calse As Currency
Dim q_sum_total As Currency
Dim q_stock As Currency
wsfile = ""
Pantalla.Enabled = False
DoEvents
'FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
PRO_REPORTE (1)
WCIA1 = ""
WCIA2 = ""
WCIA3 = ""
WCIA4 = ""

If LK_EMP <> "3AA" Then
 WCIA1 = LK_CODCIA
 GoTo OTRO
End If

For fila = 0 To LISCIA.ListCount - 1
LISCIA.ListIndex = fila
If LISCIA.Selected(fila) Then
    If Trim(WCIA1) = "" Then
     WCIA1 = Left(LISCIA.Text, 2)
    ElseIf Trim(WCIA2) = "" Then
     WCIA2 = Left(LISCIA.Text, 2)
    ElseIf Trim(WCIA3) = "" Then
     WCIA3 = Left(LISCIA.Text, 2)
    ElseIf Trim(WCIA4) = "" Then
     WCIA4 = Left(LISCIA.Text, 2)
    End If
End If
Next fila
If Trim(WCIA1) = "" And Trim(WCIA2) = "" And Trim(WCIA3) = "" And Trim(WCIA4) = "" Then
  For fila = 0 To LISCIA.ListCount - 1
    LISCIA.ListIndex = fila
    If fila = 0 Then
       WCIA1 = Left(LISCIA.Text, 2)
    End If
    If fila = 1 Then
       WCIA2 = Left(LISCIA.Text, 2)
    End If
    If fila = 2 Then
       WCIA3 = Left(LISCIA.Text, 2)
    End If
    If fila = 3 Then
       WCIA4 = Left(LISCIA.Text, 2)
    End If
  Next fila
End If
OTRO:


If Trim(ART_ARTICULO) <> "" And Trim(ART_CLASES) = "" Then
   pub_cadena = "SELECT PRE_EQUIV, ART_KEY, ART_ALTERNO, ART_NOMBRE,ART_LINEA, ARM_STOCK, ARM_COSPRO, ARM_SALDO_S FROM ARTI, ARTICULO,PRECIOS, TABLAS  WHERE (PRE_CODCIA = ART_CODCIA) AND (PRE_CODART = ART_KEY) AND (PRE_FLAG_UNIDAD = 'A') AND (ART_LINEA = TAB_NUMTAB) AND (ART_CODCIA = TAB_CODCIA) AND (TAB_TIPREG = 131) AND (ARM_CODART = ART_KEY) AND (ARM_CODCIA = ART_CODCIA) and (ARM_STOCK <> 0 OR ARM_INGRESOS <> 0 OR ARM_SALIDAS <> 0) and art_key<>0  AND ART_CODCIA = ? AND ART_KEY=  " & ART_ARTICULO & "   "
Else
   pub_cadena = "SELECT PRE_EQUIV,ART_KEY, ART_ALTERNO, ART_NOMBRE,ART_FAMILIA, ARM_STOCK, ARM_COSPRO, ARM_SALDO_S FROM ARTI, ARTICULO, TABLAS, PRECIOS WHERE (PRE_CODCIA = ART_CODCIA) AND (PRE_CODART = ART_KEY) AND (PRE_FLAG_UNIDAD = 'A') AND (ART_FAMILIA = TAB_NUMTAB) AND (ART_CODCIA = TAB_CODCIA) AND (TAB_TIPREG = 122) AND (ARM_CODART = ART_KEY) AND (ARM_CODCIA = ART_CODCIA) and ARM_SALDO_S <> 0 and art_key<>0  AND ART_CODCIA = ?  "
   If ART_CLASES <> "" Then pub_cadena = pub_cadena & " AND " & ART_CLASES
End If

pub_cadena = pub_cadena & " ORDER BY TAB_CODART, ART_NOMBRE "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

                                                                                                                                                                                                                      
pub_cadena = "SELECT FAR_FECHA_COMPRA, FAR_COSPRO_SUP, FAR_CANTIDAD, FAR_SIGNO_ARM, FAR_STOCK, FAR_COSPRO FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA_COMPRA < ? AND FAR_CODART = ? and far_estado <>'E' and far_transito <>'P' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA,  FAR_NUMOPER asc"
pub_cadena = "SELECT SUM(FAR_CANTIDAD * FAR_SIGNO_ARM)AS TOT, FAR_CODART FROM FACART WHERE FAR_CODCIA = ? AND (FAR_FECHA_COMPRA >= ? and FAR_FECHA_COMPRA <= ?) AND FAR_CODART = ? and far_estado <> 'E' and far_transito <>'P' GROUP BY FAR_CODART "
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
PS_REP03(0) = 0
PS_REP03(1) = LK_FECHA_DIA
PS_REP03(2) = LK_FECHA_DIA
PS_REP03(3) = 0
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT FAR_DESCTO,FAR_FLETE,FAR_EQUIV,FAR_COSPRO_SUP,FAR_COSPRO_ANT, FAR_TIPMOV, FAR_CODCIA, FAR_PRECIO, FAR_PRECIO_NETO,FAR_COSPRO, FAR_SUBTRA, FAR_FBG, FAR_NUMSER, FAR_NUMFAC,FAR_NUMSER_C, FAR_NUMFAC_C, FAR_FECHA_COMPRA, FAR_CANTIDAD, FAR_SIGNO_ARM, FAR_COSPRO, FAR_CODART , FAR_TIPO_CAMBIO, FAR_MONEDA, FAR_STOCK  FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ? AND FAR_CODART = ?  and far_estado<>'E' and far_transito <>'P' ORDER BY FAR_CODART, FAR_FECHA_COMPRA, FAR_NUMOPER asc "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = LK_FECHA_DIA
PS_REP01(2) = LK_FECHA_DIA
PS_REP01(3) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
DoEvents
Dim wsFECHA1, wsFECHA2
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If

ws_clave = PUB_CLAVE
GoSub WEXCEL
'FrmImp2.ProgBar.Visible = True
DoEvents
'xl.Worksheets(1).Activate
'GoSub LETRAS

'xl.Cells(1, 1) = Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia")))
xl.Cells(4, 2) = "kardex del : " & txtCampo1.Text & "  Al   " & txtCampo2.Text
f1 = 8 'Fila Inicial
PS_REP02(0) = WCIA1  ''LK_CODCIA

llave_rep02.Requery
If llave_rep02.RowCount <> 0 Then
 RCRYSTAL.ProgBar.Min = 0
 RCRYSTAL.ProgBar.Value = 0
 RCRYSTAL.ProgBar.Max = llave_rep02.RowCount
End If

RCRYSTAL.lblproceso.Visible = True
RCRYSTAL.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
CHE_KARDEX = 0
DoEvents
wtotal = 0
WD_COSPRO = 0
acu_saldo = 0
WCOSPRO_SUP = 0
TOTAL_CLASE = 0
If Not llave_rep02.EOF = True Then WW_LINEA = -1 ''llave_rep02!art_linea

Do Until llave_rep02.EOF
    ww_codcia = WCIA1 ''LK_CODCIA
    RCRYSTAL.ProgBar.Value = RCRYSTAL.ProgBar.Value + 1
    WCONCIA = 0
COME_BACK:
      WCONCIA = WCONCIA + 1
        If ww_codcia <> "01" Then
           PSART_LLAVE_ALT(0) = llave_rep02!ART_ALTERNO
           PSART_LLAVE_ALT(1) = ww_codcia
           art_llave_alt.Requery
           If art_llave_alt.EOF Then GoTo ABAJO
           PUB_KEY = art_llave_alt!ART_KEY
           PUB_CONCEPTO = art_llave_alt!art_nombre
        Else
           PUB_KEY = llave_rep02!ART_KEY
           PUB_CONCEPTO = llave_rep02!art_nombre
        End If
        f1 = f1 + 1
        PS_REP03(0) = ww_codcia
        PS_REP03(1) = Format(txtCampo1.Text, "dd/mm/yyyy")
        PS_REP03(2) = LK_FECHA_DIA
        PS_REP03(3) = PUB_KEY
        llave_rep03.Requery
        
        WCOSPRO_SUP = 0
        PUB_IMPORTE = 0
        PUB_IMPORTE_AMORT = 0
        CHE_KARDEX = 0
        WCOSPRO_SUP = 0
        PUB_IMPORTE = 0 ' llave_rep03!FAR_STOCK
        If llave_rep03.EOF Then
           PUB_IMPORTE = val(llave_rep02!arm_stock)
        Else
          If (val(llave_rep03!TOT) - val(llave_rep02!arm_stock)) < 0 Then
            PUB_IMPORTE = Abs(val(llave_rep03!TOT) - val(llave_rep02!arm_stock))
          Else
            PUB_IMPORTE = (val(llave_rep03!TOT) - val(llave_rep02!arm_stock)) * -1
          End If
        End If
        
        If WW_LINEA <> llave_rep02!art_familia Then
                If f1 <> 6 Then
                 wranF = "G" & f1 & ":G" & f1
                 xl.Range(wranF).Font.Bold = True
                 xl.Range(wranF).Font.Name = "Arial"
                 xl.Range(wranF).Font.Size = 9
                 xl.Worksheets(1).rows(f1).RowHeight = 11
                ' xl.Cells(f1, 8) = "Total Division: "
                ' xl.Cells(f1, 10) = q_stock
                ' xl.Cells(f1, 12) = q_stock_val
                 q_stock = 0
                 q_stock_val = 0
                End If
                PUB_TIPREG = 122
                PUB_NUMTAB = llave_rep02!art_familia
                PUB_CODCIA = ww_codcia
                SQ_OPER = 1
                LEER_TAB_LLAVE
                If f1 <> 6 Then f1 = f1 + 1
                wranF = "A" & f1 & ":A" & f1
                xl.Range(wranF).Font.Bold = True
                xl.Range(wranF).Font.Name = "Arial"
                xl.Range(wranF).Font.Size = 11
                xl.Worksheets(1).rows(f1).RowHeight = 12
'                If tab_llave.EOF Then
'                   xl.Cells(f1, 1) = "DIVISION: "
'                Else
'                  xl.Cells(f1, 1) = "DIVISION: " & Trim(tab_llave!tab_NOMLARGO)
'                End If
'                WW_LINEA = llave_rep02!art_familia
        End If
        PS_REP01(0) = ww_codcia
        PS_REP01(1) = Format(txtCampo1.Text, "dd/mm/yyyy")
        PS_REP01(2) = Format(txtCampo2.Text, "dd/mm/yyyy")
        PS_REP01(3) = PUB_KEY
        llave_rep01.Requery
        
        pu_codcia = ww_codcia
        SQ_OPER = 1
        PUB_SECUEN = 0
        PUB_CODART = PUB_KEY
        LEER_PRE_LLAVE
        f1 = f1 + 1
        xl.Cells(f1, 1) = "'" & "03" '& " " & ww_codcia
        xl.Cells(f1, 2) = llave_rep02!ART_ALTERNO & "-" & Trim(PUB_CONCEPTO)
        'xl.Cells(f1, 6) = Trim("03") 'Trim(PUB_CONCEPTO)"
        
       ' If Trim(pre_llave!pre_unidad) = "UNIDAD" Then
       '  xl.Cells(f1, 5) = "07" & " " & "UNIDAD"
       ' ElseIf Trim(pre_llave!pre_unidad) = "KG" Then
       '  xl.Cells(f1, 5) = "01"
       ' ElseIf Trim(pre_llave!pre_unidad) = "LITRO" Then
       '  xl.Cells(f1, 5) = "08"
       ' Else
         xl.Cells(f1, 5) = Trim(pre_llave!pre_unidad) '"99"
       ' End If
        
        f1 = f1 + 1
        xl.Cells(f1, 1) = "Saldo Inicial "
        'xl.Cells(f1, 11) = Format(WCOSPRO_SUP, "0.00") 'Val(llave_rep01!FAR_COSPRO)
'        PUB_IMPORTE = (llave_rep01!FAR_STOCK + ((llave_rep01!far_SIGNO_aRM * llave_rep01!FAR_CANTIDAD) * -1))
        PUB_IMPORTE = Format(PUB_IMPORTE / llave_rep02!PRE_EQUIV, "0.00")
        xl.Cells(f1, 8) = PUB_IMPORTE
        PUB_IMPORTE_AMORT = PUB_IMPORTE * WCOSPRO_SUP
       ' xl.Cells(f1, 12) = val(PUB_IMPORTE_AMORT)
        If Not llave_rep01.EOF Then
        End If
        
       ' PUB_IMPORTE_AMORT = 0
       ' PUB_IMPORTE = 0
        CHE_KARDEX = PUB_IMPORTE
        xx_ingreso = 0
        xx_salida = 0
        acu_val_ingresos = 0
        acu_val_salidas = 0
        flag_xx = 0
       ' WCOSPRO_SUP = 0
'       xl.Application.Visible = True
       ' If Not llave_rep01.EOF Then
       '   WCOSPRO_SUP = Nulo_Valor0(llave_rep01!FAR_COSPRO_SUP)
       ' End If
        Do Until llave_rep01.EOF
           If llave_rep01!FAR_fecha_compra > CDate(txtCampo2.Text) Then Exit Do
            If flag_xx = 0 Then
              FF1 = f1
           End If
           f1 = f1 + 1
           WTC = 1
           If llave_rep01!FAR_MONEDA = "D" Then
              WTC = JALAR(llave_rep01!FAR_fecha_compra)
              WS_PRECIO = Format(llave_rep01!FAR_PRECIO * WTC, "0.00")
           Else
              WS_PRECIO = llave_rep01!FAR_PRECIO
           End If
'           xl.Application.Visible = True
            WCAT = Format(llave_rep01!far_cantidad / llave_rep02!PRE_EQUIV, "0.00")
           If llave_rep01!far_signo_arm = 1 Then
              ww_ingreso = WCAT * WS_PRECIO
              ww_ingreso = (WTC * WCAT * llave_rep01!FAR_PRECIO / llave_rep01!FAR_equiv) + llave_rep01!FAR_FLETE - redondea((llave_rep01!FAR_DESCTO * WTC))
              If llave_rep01!FAR_TIPMOV = 20 Then ww_ingreso = ww_ingreso ' llave_rep01!far_precio_neto
              ww_ingreso = ww_ingreso
              acu_val_ingresos = acu_val_ingresos + ww_ingreso
              xl.Cells(f1, 6) = WCAT
              xx_ingreso = xx_ingreso + WCAT
          '    xl.Cells(f1, 8) = val(ww_ingreso)
              CHE_KARDEX = CHE_KARDEX + WCAT
              
           End If
           If llave_rep01!far_signo_arm = -1 Then
              ww_salida = WCAT * llave_rep01!FAR_COSPRO
              acu_val_salidas = acu_val_salidas + ww_salida
              xl.Cells(f1, 7) = WCAT
              xx_salida = xx_salida + WCAT
           '   xl.Cells(f1, 9) = val(ww_salida)
              CHE_KARDEX = CHE_KARDEX - WCAT
           End If
           
           If llave_rep01!FAR_TIPMOV = 20 Then
              ww_concepto = "01" 'llave_rep01!far_fbg & " " & Trim(llave_rep01!far_numser) & "-" & llave_rep01!far_numfac
           ElseIf llave_rep01!FAR_TIPMOV = 5 Then
              ww_concepto = "99" 'Trim(llave_rep01!far_numser) & "-" & llave_rep01!far_numfac
           ElseIf llave_rep01!FAR_TIPMOV = 6 Then
           ww_concepto = "99" 'Trim(llave_rep01!far_numser) & "-" & llave_rep01!far_numfac
           End If
           xl.Cells(f1, 2) = "'" & Trim(ww_concepto)
           xl.Cells(f1, 1) = "'" & Format(llave_rep01!FAR_fecha_compra, "dd/mm/yyyy")
           If llave_rep01!FAR_TIPMOV = 20 Then
              xl.Cells(f1, 3) = llave_rep01!FAR_NUMSER_C
              xl.Cells(f1, 4) = llave_rep01!FAR_NUMFAC_C
           Else
              xl.Cells(f1, 3) = llave_rep01!far_numser
              xl.Cells(f1, 4) = llave_rep01!far_numfac
           End If
           
           If llave_rep01!FAR_TIPMOV = 20 Then
              ww_tipo = "02" 'llave_rep01!far_fbg & " " & Trim(llave_rep01!far_numser) & "-" & llave_rep01!far_numfac
           ElseIf llave_rep01!FAR_TIPMOV = 5 Then
              ww_tipo = "10 " 'Trim(llave_rep01!far_numser) & "-" & llave_rep01!far_numfac
           ElseIf llave_rep01!FAR_TIPMOV = 6 Then
           ww_tipo = "11" 'Trim(llave_rep01!far_numser) & "-" & llave_rep01!far_numfac
           End If
           xl.Cells(f1, 5) = "'" & Trim(ww_tipo)
           
          ' xl.Cells(f1, 7) = WS_PRECIO
           If flag_xx = 0 Then
            flag_xx = 1
            PUB_COSPRO = llave_rep01!FAR_COSPRO
          '  xl.Cells(f1, 11) = llave_rep01!FAR_COSPRO
           Else
            PUB_COSPRO = llave_rep01!FAR_COSPRO
         '   xl.Cells(f1, 11) = val(llave_rep01!FAR_COSPRO)
           End If
           PUB_IMPORTE = CHE_KARDEX 'llave_rep01!FAR_STOCK
           If CHE_KARDEX <> PUB_IMPORTE Then
'              MsgBox "Hacer el Proceso del Costeo del Articulo, Codigo:  " & llave_rep01!FAR_fecha_compra & " " & llave_rep02!art_alterno
           End If
           PUB_IMPORTE_AMORT = PUB_IMPORTE * PUB_COSPRO
'           If Val(PUB_IMPORTE) = 0 Then Stop
           xl.Cells(f1, 8) = val(PUB_IMPORTE)
        '   xl.Cells(f1, 12) = val(PUB_IMPORTE_AMORT)
    
           llave_rep01.MoveNext
           
           
        Loop
        'If flag_xx = 0 Then GoTo ABAJO
       'xl.Application.Visible = True
        f1 = f1 + 1
        xl.Cells(f1, 1) = "Stock al : " & txtCampo2.Text
        xl.Cells(f1, 8) = PUB_IMPORTE
       ' xl.Cells(f1, 12) = PUB_IMPORTE_AMORT
        wtotal = wtotal + PUB_IMPORTE_AMORT
       ' xl.Cells(f1, 9) = acu_val_salidas
       ' xl.Cells(f1, 8) = acu_val_ingresos
        xl.Cells(f1, 6) = xx_ingreso
        xl.Cells(f1, 7) = xx_salida
        q_stock = q_stock + PUB_IMPORTE
        TOTAL_CLASE = TOTAL_CLASE + PUB_IMPORTE
        q_stock_val = q_stock_val + PUB_IMPORTE_AMORT
        TOTAL_CLASE_VAL = TOTAL_CLASE_VAL + PUB_IMPORTE_AMORT
        

        
ABAJO:
 If LK_EMP = "3AA" Then
    If WCONCIA = 1 Then
        ww_codcia = WCIA2
        GoTo COME_BACK
    End If
    If WCONCIA = 2 Then
       ww_codcia = WCIA3
       GoTo COME_BACK
    End If
    If WCONCIA = 3 Then
       ww_codcia = WCIA4
       GoTo COME_BACK
    End If
 End If
 'WW_LINEA = llave_rep02!art_linea
llave_rep02.MoveNext
Loop
   f1 = f1 + 1
   wranF = "G" & f1 & ":G" & f1
   xl.Range(wranF).Font.Bold = True
   xl.Range(wranF).Font.Name = "Arial"
   xl.Range(wranF).Font.Size = 9
   xl.Worksheets(1).rows(f1).RowHeight = 11
  ' xl.Cells(f1, 8) = "Total Division: "
   xl.Cells(f1, 6) = q_stock
  ' xl.Cells(f1, 12) = q_stock_val
   f1 = f1 + 1
   wranF = "H" & f1 & ":H" & f1
   xl.Range(wranF).Font.Bold = True
   xl.Range(wranF).Font.Name = "Arial"
   xl.Range(wranF).Font.Size = 10
   xl.Worksheets(1).rows(f1).RowHeight = 12
  ' xl.Cells(f1, 8) = "TOTAL GENERAL = "
   xl.Cells(f1, 6) = Format(TOTAL_CLASE, "#,##0.000")
 '  xl.Cells(f1, 12) = Format(TOTAL_CLASE_VAL, "#,##0.000")

  RCRYSTAL.lblproceso.Caption = "Procesando . . .  un Momento ."
  'xl.Application.Worksheets("HOJA DE COMISIONES x VENDEDOR").Range(wran1).Sort Key1:=xl.Application.Worksheets("HOJA DE COMISIONES x VENDEDOR").Range("O7")
  RCRYSTAL.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
 ' xl.Cells(2, 2) = "INFORME DE COMISIONES x VENDEDOR"
  DoEvents
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect PUB_CLAVE
  xl.APPLICATION.Visible = True
  DoEvents
  RCRYSTAL.lblproceso.Visible = False
  RCRYSTAL.ProgBar.Visible = False
  Set xl = Nothing
    Screen.MousePointer = 0
  ProgBar.Visible = False
  lblproceso.Visible = False
  Pantalla.Enabled = True
  cmdCerrar.Enabled = True
  ''Unload RCRYSTAL
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
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblproceso.Caption = "Abriendo , Archivo Comisiones.xls . . . "
  DoEvents
  xl.Workbooks.Open Left(Trim(PUB_RUTA_OTRO), 1) & ":\ADMIN\STANDAR\KARDEX_CLASES.xls", 0, True, 4

Return



Exit Sub
CANCELA:
  RCRYSTAL.Pantalla.Enabled = True
  RCRYSTAL.Pantalla.Caption = "Por &Pantalla"
  RCRYSTAL.lblproceso.Visible = False
  RCRYSTAL.ProgBar.Visible = False
  Pantalla.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
  
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
OJO:
If Err.Number = 70 Then
  MsgBox "Hoja de Calculo : " & wsfile1 & "  esta Abierta debe cerrar para Procesar Nuevamente ", 48, Pub_Titulo
  GoTo CANCELA
End If
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 xl.APPLICATION.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
'' Unload FrmImp2
Exit Sub
End Sub

Public Sub RESU_KARDEX()
'On Error GoTo FINTODO
Dim CHE_KARDEX As Currency
Dim WSUMCIA As Integer
Dim WCONCIA As Integer
Dim wsalmacenes As String
Dim WTC As Currency
Dim wtotal As Currency
Dim WCOSPRO_SUP As Currency
Dim TOTAL_CLASE As Currency
Dim WSUMA_CALSE As Currency
Dim WCODIGO As String
Dim wnombre As String
Dim wunidad As String
Dim WCIA1 As String * 2
Dim WCIA2 As String * 2
Dim WCIA3 As String * 2
Dim WCIA4 As String * 2
Dim WSCODART As Currency
Dim flag_xx As Integer
Dim ww_concepto As String
Dim ww_codcia As String * 2
Dim WS_PRECIO As Currency
Dim WW_LINEA, i
Dim ws_clave As String
Dim FF1 As Integer
Dim LETRAS(24) As String * 1
Dim WSFECHA As Date
Dim acu_val_ingresos As Currency
Dim acu_val_salidas As Currency
Dim xx_ingreso As Currency
Dim xx_salida As Currency
Dim ww_ingreso As Currency
Dim ww_salida As Currency
Dim acu_cant_dia As Currency
Dim acu_saldo As Currency
Dim acu_stock As Currency
Dim wsfile As String
Dim walterno As String * 10
Dim wdnombre As String
Dim WD_COSPRO As Currency

Dim INICIAL As Currency
Dim COMPRA As Currency
Dim VENTA As Currency
Dim AJSAL As Currency
Dim AJING As Currency
Dim ENVIO As Currency
Dim RECEP As Currency
Dim CAMBIOI As Currency
Dim CAMBIOS As Currency

walterno = ""
wsfile = ""
Pantalla.Enabled = False
DoEvents
'FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
If Not SON_FECHAS Then
  Exit Sub
End If
PRO_REPORTE (1)
WCIA1 = ""
WCIA2 = ""
WCIA3 = ""
WCIA4 = ""
WSUMCIA = 0
wsalmacenes = ""

If LK_EMP <> "3AA" Then
 WCIA1 = LK_CODCIA
 GoTo OTRO
End If
For fila = 0 To LISCIA.ListCount - 1
 LISCIA.ListIndex = fila
 If LISCIA.Selected(fila) Then
    PSPAR_MULTI(0) = Left(LISCIA.Text, 2)
    par_multi.Requery
    wsalmacenes = wsalmacenes + Trim(par_multi!par_nombre_corto) & " - "
 End If
Next fila
If wsalmacenes <> "" Then
    wsalmacenes = Mid(wsalmacenes, 1, Len(wsalmacenes) - 3)
End If
        
        


For fila = 0 To LISCIA.ListCount - 1
LISCIA.ListIndex = fila
If LISCIA.Selected(fila) Then
    If Trim(WCIA1) = "" Then
     WCIA1 = Left(LISCIA.Text, 2)
    ElseIf Trim(WCIA2) = "" Then
     WCIA2 = Left(LISCIA.Text, 2)
    ElseIf Trim(WCIA3) = "" Then
     WCIA3 = Left(LISCIA.Text, 2)
    ElseIf Trim(WCIA4) = "" Then
     WCIA4 = Left(LISCIA.Text, 2)
    End If
End If
Next fila
WSUMCIA = 0
If Trim(WCIA1) = "" And Trim(WCIA2) = "" And Trim(WCIA3) = "" And Trim(WCIA4) = "" Then
  For fila = 0 To LISCIA.ListCount - 1
    LISCIA.ListIndex = fila
    If fila = 0 Then
       WCIA1 = Left(LISCIA.Text, 2)
       WSUMCIA = WSUMCIA + 1
    End If
    If fila = 1 Then
       WCIA2 = Left(LISCIA.Text, 2)
       WSUMCIA = WSUMCIA + 1
    End If
    If fila = 2 Then
       WCIA3 = Left(LISCIA.Text, 2)
       WSUMCIA = WSUMCIA + 1
    End If
    If fila = 3 Then
       WCIA4 = Left(LISCIA.Text, 2)
       WSUMCIA = WSUMCIA + 1
    End If
  Next fila
End If
OTRO:



If Trim(ART_ARTICULO) <> "" And Trim(ART_CLASES) = "" And Trim(ART_LINEAS) = "" Then
 '  pub_cadena = "SELECT ART_KEY, ART_ALTERNO, ART_NOMBRE,ART_LINEA, ARM_STOCK, ARM_COSPRO, ARM_SALDO_S FROM ARTI, ARTICULO  WHERE (ARM_CODART = ART_KEY) AND (ARM_CODCIA = ART_CODCIA)  and art_key<>0  AND ART_CODCIA = ? AND ART_KEY=  " & ART_ARTICULO
   pub_cadena = "SELECT ART_KEY, ART_ALTERNO, ART_NOMBRE,ART_LINEA, ARM_STOCK, ARM_COSPRO, ARM_SALDO_S FROM ARTI, ARTICULO  WHERE (ARM_CODART = ART_KEY) AND (ARM_CODCIA = ART_CODCIA)  and art_key <> 0  AND ART_CODCIA = ? AND ART_KEY=  " & ART_ARTICULO
Else
   pub_cadena = "SELECT ART_KEY, ART_ALTERNO, ART_NOMBRE,ART_LINEA, ARM_STOCK, ARM_COSPRO, ARM_SALDO_S FROM ARTI, ARTICULO  WHERE (ARM_CODART = ART_KEY) AND (ARM_CODCIA = ART_CODCIA)  and art_key <> 0  AND ART_CODCIA = ? "
   If ART_CLASES <> "" Then pub_cadena = pub_cadena & " AND " & ART_CLASES
   If ART_LINEAS <> "" Then pub_cadena = pub_cadena & " AND " & ART_LINEAS
End If
pub_cadena = pub_cadena & " ORDER BY ART_LINEA, ART_ALTERNO "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT FAR_DESCTO, FAR_FLETE,FAR_EQUIV,FAR_COSPRO_SUP,FAR_COSPRO_ANT, FAR_TIPMOV, FAR_CODCIA, FAR_PRECIO, FAR_PRECIO_NETO,FAR_COSPRO, FAR_SUBTRA, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_FECHA_COMPRA, FAR_CANTIDAD, FAR_SIGNO_ARM, FAR_COSPRO, FAR_CODART , FAR_TIPO_CAMBIO, FAR_MONEDA, FAR_STOCK  FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ? AND FAR_CODART = ?  and far_estado<>'E' ORDER BY FAR_CODART, FAR_FECHA_COMPRA,FAR_SIGNO_ARM DESC, FAR_NUMOPER2 "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = LK_FECHA_DIA
PS_REP01(2) = LK_FECHA_DIA
PS_REP01(3) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

'pub_cadena = "SELECT FAR_STOCK,FAR_COSPRO FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA_COMPRA <= ? AND FAR_CODART = ? and FAR_ESTADO <>'E' ORDER BY FAR_FECHA_COMPRA,FAR_SIGNO_ARM DESC, FAR_NUMOPER2"
pub_cadena = "SELECT FAR_FECHA_COMPRA, FAR_COSPRO_SUP, FAR_CANTIDAD, FAR_SIGNO_ARM, FAR_STOCK, FAR_COSPRO FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA_COMPRA < ? AND FAR_CODART = ? and far_estado <>'E' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA, FAR_SIGNO_ARM DESC , FAR_NUMOPER2"
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
PS_REP03(0) = 0
PS_REP03(1) = LK_FECHA_DIA
PS_REP03(2) = 0
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


DoEvents
Dim wsFECHA1, wsFECHA2
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If

ws_clave = PUB_CLAVE
GoSub WEXCEL
'FrmImp2.ProgBar.Visible = True
DoEvents
'xl.Worksheets(1).Activate
'GoSub LETRAS

xl.Cells(1, 1) = Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia")))
xl.Cells(2, 1) = "RESUMEN KARDEX VALORIZADO " + wsalmacenes
xl.Cells(4, 3) = "kardex del : " & txtCampo1.Text & "  Al   " & txtCampo2.Text
f1 = 6  'Fila Inicial
PS_REP02(0) = WCIA1  ''LK_CODCIA

llave_rep02.Requery
If llave_rep02.RowCount <> 0 Then
 RCRYSTAL.ProgBar.Min = 0
 RCRYSTAL.ProgBar.Value = 0
 RCRYSTAL.ProgBar.Max = llave_rep02.RowCount
End If

RCRYSTAL.lblproceso.Visible = True
RCRYSTAL.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
wtotal = 0
WD_COSPRO = 0
acu_saldo = 0
WCOSPRO_SUP = 0
If Not llave_rep02.EOF = True Then WW_LINEA = -1 ''llave_rep02!art_linea
 'xl.Application.Visible = True
 WCONCIA = 0
Do Until llave_rep02.EOF
    ww_codcia = WCIA1 ''LK_CODCIA
    RCRYSTAL.ProgBar.Value = RCRYSTAL.ProgBar.Value + 1
     WCONCIA = 0
COME_BACK:
        WCONCIA = WCONCIA + 1
        If ww_codcia <> "01" Then
           PSART_LLAVE_ALT(0) = llave_rep02!ART_ALTERNO
           PSART_LLAVE_ALT(1) = ww_codcia
           art_llave_alt.Requery
           If art_llave_alt.EOF Then GoTo ABAJO
           PUB_KEY = art_llave_alt!ART_KEY
           PUB_CONCEPTO = art_llave_alt!art_nombre
        Else
           PUB_KEY = llave_rep02!ART_KEY
           PUB_CONCEPTO = llave_rep02!art_nombre
        End If
 '       F1 = F1 + 1
        PS_REP03(0) = ww_codcia
        PS_REP03(1) = Format(txtCampo1.Text, "dd/mm/yyyy")
        PS_REP03(2) = PUB_KEY
        llave_rep03.Requery
        llave_rep03.MoveLast
        WCOSPRO_SUP = 0
        PUB_IMPORTE = 0
        PUB_IMPORTE_AMORT = 0
        CHE_KARDEX = 0
        If Not llave_rep03.EOF Then
  '          If llave_rep03!FAR_FECHA_COMPRA = Format(txtCampo2.Text, "dd/mm/yyyy") Then
             PUB_COSPRO = llave_rep03!FAR_COSPRO
             PUB_IMPORTE = (llave_rep03!FAR_STOCK) ' + ((llave_rep03!far_SIGNO_aRM * llave_rep03!FAR_CANTIDAD) * -1))
             PUB_IMPORTE_AMORT = PUB_IMPORTE * PUB_COSPRO
   '         End If
        End If
        CHE_KARDEX = PUB_IMPORTE
        If WW_LINEA <> llave_rep02!art_linea Then
                 If WW_LINEA <> -1 Then
                  f1 = f1 + 1
                  xl.Cells(f1, 9) = "TOTAL = "
                  xl.Cells(f1, 11) = Format(WSUMA_CALSE, "#,##0.000")
                  TOTAL_CLASE = TOTAL_CLASE + WSUMA_CALSE
                End If
                PUB_TIPREG = 131
                PUB_NUMTAB = llave_rep02!art_linea
                PUB_CODCIA = ww_codcia
                SQ_OPER = 1
                LEER_TAB_LLAVE
                f1 = f1 + 1
                wranF = "A" & f1 & ":A" & f1
                xl.Range(wranF).Font.Bold = True
                xl.Range(wranF).Font.Name = "Arial"
                xl.Range(wranF).Font.Size = 12
                If tab_llave.EOF Then
                   xl.Cells(f1, 1) = "CLASE: "
                Else
                  xl.Cells(f1, 1) = "CLASE: " & Trim(tab_llave!tab_NOMLARGO)
                End If
                WW_LINEA = llave_rep02!art_linea
                WSUMA_CALSE = 0
        End If
        
        PS_REP01(0) = ww_codcia
        PS_REP01(1) = Format(txtCampo1.Text, "dd/mm/yyyy")
        PS_REP01(2) = Format(txtCampo2.Text, "dd/mm/yyyy")
        PS_REP01(3) = PUB_KEY
        llave_rep01.Requery
        
        
        walterno = llave_rep02!ART_ALTERNO
        wdnombre = llave_rep02!art_nombre
        pu_codcia = ww_codcia
        'SQ_OPER = 1
        'PUB_SECUEN = 0
        PUB_CODART = PUB_KEY
        'LEER_PRE_LLAVE
        
        xx_ingreso = 0
        xx_salida = 0
        acu_val_ingresos = 0
        acu_val_salidas = 0
        flag_xx = 0
        
        
        Do Until llave_rep01.EOF
           If llave_rep01!FAR_fecha_compra > CDate(txtCampo2.Text) Then Exit Do
           WTC = 1
           If llave_rep01!FAR_MONEDA = "D" Then
             WTC = JALAR(llave_rep01!FAR_fecha_compra)
              WS_PRECIO = Format(llave_rep01!FAR_PRECIO * WTC, "0.0000")
           Else
              WS_PRECIO = llave_rep01!FAR_PRECIO
           End If
'            xl.Application.Visible = True
           If llave_rep01!far_signo_arm = 1 Then
              ww_ingreso = llave_rep01!far_cantidad * WS_PRECIO
              ww_ingreso = (WTC * llave_rep01!far_cantidad * llave_rep01!FAR_PRECIO / llave_rep01!FAR_equiv) + llave_rep01!FAR_FLETE - redondea(llave_rep01!FAR_DESCTO * WTC)
              If llave_rep01!FAR_TIPMOV = 20 Then ww_ingreso = ww_ingreso ' llave_rep01!far_precio_neto
              ww_ingreso = ww_ingreso
              acu_val_ingresos = acu_val_ingresos + ww_ingreso
              xx_ingreso = xx_ingreso + llave_rep01!far_cantidad
              CHE_KARDEX = CHE_KARDEX + val(llave_rep01!far_cantidad)
           End If
           If llave_rep01!far_signo_arm = -1 Then
              ww_salida = llave_rep01!far_cantidad * llave_rep01!FAR_COSPRO
              acu_val_salidas = acu_val_salidas + ww_salida
              xx_salida = xx_salida + llave_rep01!far_cantidad
              CHE_KARDEX = CHE_KARDEX - val(llave_rep01!far_cantidad)
           End If
           PUB_COSPRO = llave_rep01!FAR_COSPRO
           PUB_IMPORTE = llave_rep01!FAR_STOCK
           If CHE_KARDEX <> PUB_IMPORTE Then
            '  MsgBox "Hacer el Proceso del Costeo del Articulo, Codigo:  " & llave_rep01!FAR_fecha_compra & " " & llave_rep02!ART_ALTERNO & "-" & llave_rep02!art_nombre
              'CHE_KARDEX <> PUB_IMPORTE
           End If
           PUB_IMPORTE_AMORT = PUB_IMPORTE * PUB_COSPRO
           llave_rep01.MoveNext
        Loop
       'If flag_xx = 0 Then GoTo ABAJO
       'xl.Application.Visible = True
        f1 = f1 + 1
        xl.Cells(f1, 1) = walterno & Format(ww_codcia, "00")
        xl.Cells(f1, 2) = wdnombre
        xl.Cells(f1, 3) = "" 'wunidad
        xl.Cells(f1, 9) = PUB_IMPORTE
        xl.Cells(f1, 10) = PUB_COSPRO
        xl.Cells(f1, 11) = PUB_IMPORTE_AMORT
        wtotal = wtotal + PUB_IMPORTE_AMORT
        xl.Cells(f1, 8) = acu_val_salidas
        xl.Cells(f1, 7) = acu_val_ingresos
        xl.Cells(f1, 4) = xx_ingreso
        xl.Cells(f1, 5) = xx_salida
        WSUMA_CALSE = WSUMA_CALSE + redondea(PUB_IMPORTE_AMORT)
        
ABAJO:
  If LK_EMP = "3AA" Then
    If WCONCIA = 1 Then
        ww_codcia = WCIA2
        GoTo COME_BACK
    End If
    If WCONCIA = 2 Then
       ww_codcia = WCIA3
       GoTo COME_BACK
    End If
    If WCONCIA = 3 Then
       ww_codcia = WCIA4
       GoTo COME_BACK
    End If
 End If
''WW_LINEA = llave_rep02!art_linea
llave_rep02.MoveNext
Loop
 'MsgBox WTOTAL
   f1 = f1 + 1
   xl.Cells(f1, 9) = "TOTAL = "
   xl.Cells(f1, 11) = Format(WSUMA_CALSE, "#,##0.000")
   TOTAL_CLASE = TOTAL_CLASE + WSUMA_CALSE
   f1 = f1 + 1
   xl.Cells(f1, 9) = "TOTAL GENERAL = "
   xl.Cells(f1, 11) = Format(TOTAL_CLASE, "#,##0.000")
   

  RCRYSTAL.lblproceso.Caption = "Procesando . . .  un Momento ."
  'xl.Application.Worksheets("HOJA DE COMISIONES x VENDEDOR").Range(wran1).Sort Key1:=xl.Application.Worksheets("HOJA DE COMISIONES x VENDEDOR").Range("O7")
  RCRYSTAL.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
 ' xl.Cells(2, 2) = "INFORME DE COMISIONES x VENDEDOR"
  DoEvents
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  RCRYSTAL.lblproceso.Visible = False
  RCRYSTAL.ProgBar.Visible = False
  Set xl = Nothing
   Screen.MousePointer = 0
  ProgBar.Visible = False
  lblproceso.Visible = False
  Pantalla.Enabled = True
  cmdCerrar.Enabled = True
  ''Unload RCRYSTAL
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
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblproceso.Caption = "Abriendo , Archivo Comisiones.xls . . . "
  DoEvents
  xl.Workbooks.Open Left(Trim(PUB_RUTA_OTRO), 1) & ":\ADMIN\STANDAR\KARDEX_RESU.xls", 0, True, 4

Return



Exit Sub
CANCELA:
  RCRYSTAL.Pantalla.Enabled = True
  RCRYSTAL.Pantalla.Caption = "Por &Pantalla"
  RCRYSTAL.lblproceso.Visible = False
  RCRYSTAL.ProgBar.Visible = False
  Pantalla.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
  
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
OJO:
If Err.Number = 70 Then
  MsgBox "Hoja de Calculo : " & wsfile1 & "  esta Abierta debe cerrar para Procesar Nuevamente ", 48, Pub_Titulo
  GoTo CANCELA
End If
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 xl.APPLICATION.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
'' Unload FrmImp2
Exit Sub
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
     codi = Format(llave_rep01!VEM_CODVEN, "000")
     multiven.AddItem codi & " " & Trim(llave_rep01!VEM_NOMBRE)
     llave_rep01.MoveNext
 Loop
 multiven.Visible = True
 fraven.Visible = True
End Sub

Private Sub txtCampo2_KeyPress(KeyAscii As Integer)
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
  If loc_key > ListView3.ListItems.count Then loc_key = ListView3.ListItems.count
  GoTo posicion
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo posicion
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView3.ListItems.count Then loc_key = ListView3.ListItems.count
 GoTo posicion
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo posicion
End If
GoTo fin
posicion:
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
pu_codclie = val(txt_key.Text)
If Len(txt_key.Text) = 0 Then
   Exit Sub
End If
'fra2.Refresh
If pu_codclie <> 0 And IsNumeric(txt_key.Text) = True Then
    SQ_OPER = 1
    On Error GoTo mucho
    PUB_CODBAN = val(txt_key.Text)
    On Error GoTo 0
    pu_codcia = LK_CODCIA
    LEER_CCM_LLAVE
    If ccm_llave.EOF Then
            MsgBox "Registro ,   NO EXISTE ... "
            Azul txt_key, txt_key
            GoTo fin
    End If
    lblbanco.Caption = Trim(ccm_llave!CCM_NOMBRE)
    txt_key.Text = Trim(ccm_llave!CCM_CODBAN)
    If Pantalla.Visible And Pantalla.Enabled Then
      Pantalla.SetFocus
    End If
    ListView3.Visible = False

    Screen.MousePointer = 0
Else
   If loc_key > ListView3.ListItems.count Or loc_key = 0 Then
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
Dim VAR
If Len(txt_key.Text) = 0 Or IsNumeric(txt_key.Text) = True Then
   ListView3.Visible = False
   Exit Sub
End If
If ListView3.Visible = False And KeyCode <> 13 Or Len(txt_key.Text) = 1 Then
    If txt_key.Text = "" Then txt_key.Text = " "
    VAR = Asc(txt_key.Text)
    VAR = VAR + 1
    NADA = VAR
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    numarchi = 9
    archi = "SELECT * FROM CCMAEST WHERE  CCM_CODCIA = '" & par_llave!PAR_CIACCM & "' AND CCM_NOMBRE BETWEEN '" & txt_key.Text & "' AND  '" & VAR & "' ORDER BY CCM_NOMBRE"
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
   If loc_key + 8 > ListView3.ListItems.count Then
      ListView3.ListItems.Item(ListView3.ListItems.count).EnsureVisible
   Else
     ListView3.ListItems.Item(loc_key + 8).EnsureVisible
   End If
  End If
  Exit Sub
End If
End Sub

Private Sub ListView3_DblClick()
 loc_key = ListView3.SelectedItem.index
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
 loc_key = ListView3.SelectedItem.index
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

Public Sub REG_COMPRA_COM()
On Error GoTo FINTODO
Dim Lini As Integer
Dim Lfin As Integer
Dim qver_onlyCont As Integer
Dim CHE_IMPORTE As Currency
Dim wsigno As Currency
Dim WEMPRESA As String
Dim WCHE_TOTAL As Currency
Dim WCHE_IGV As Currency
Dim xcuenta  As Integer
Dim wCTAR2 As Currency
Dim wCTARCTA2  As String * 12
Dim fca1 As String * 1
Dim fca2 As String * 1
Dim IMP_CTA1 As Currency
Dim IMP_CTA2 As Currency
Dim WTC As Currency
Dim wOTRO As Currency
Dim wOTROCTA As String
Dim wFLETE As Currency
Dim wFLETECTA As String
Dim wCTAR As Currency
Dim wCTARCTA As String
Dim wDescto  As Currency
Dim wDesctoCTA As String
Dim LETRAS(100) As String * 2
Dim FILTRO_CTA(3) As String
Dim wsFECHA1
Dim wsFECHA2
Dim wcta1 As String
Dim wIMPORTE1 As Currency
Dim IMP_MONEDA As String * 1
Pantalla.Enabled = False
cmdCerrar.Enabled = False
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
If chepasa.Value = 1 Then
  pub_mensaje = "<Advertencia> El pase de la información es por cada Compañia. Continuar...?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then
      Azul2 txtCampo1, txtCampo1
      GoTo CANCELA
  End If
  qver_onlyCont = CHE_BLOQ_MES(1)
  If qver_onlyCont = 1 Then
      MsgBox "Periodo Activo esta Cerrado. No procede.", 48, Pub_Titulo
      chepasa.Value = 0
      GoTo CANCELA
  End If
  If (cop_llave!cop_fecha_proceso = CDate(wsFECHA1)) And (cop_llave!cop_fecha_proceso2 = CDate(wsFECHA2)) Then
  Else
      MsgBox "Usted. a marcado la opción: Pasar la Información al Periodo Contable. " & Chr(13) & Chr(13) & "Las Fechas ingresadas son distintas a la del Periodo Contable Activo. Verificar...", 48, Pub_Titulo
      Azul2 txtCampo1, txtCampo1
      GoTo CANCELA
  End If
  If qver_onlyCont = 9 Then ' hay Información en OnlyCont. Confirmar.
      pub_mensaje = "Usted. a marcado la opción: Pasar la Información al Periodo Contable. " & Chr(13) & Chr(13) & "Existe Voucher en el Periodo Contable Activo. " & Chr(13) & Chr(13) & "< Desea adicionar este Nuevo Asiento de Voucher de todas Maneras >...?"
      Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
      If Pub_Respuesta = vbNo Then
        Azul2 txtCampo1, txtCampo1
        GoTo CANCELA
      End If
  End If
End If

GoSub WEXCEL
pub_cadena = ""
'xl.Application.Visible = True
xcuenta = 0

Pantalla.Enabled = False
cmdCerrar.Enabled = False
DoEvents
RCRYSTAL.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
If opcompra(0).Value Then
  'pub_cadena = "SELECT * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ? OR ALL_CODCIA = ? ) AND (ALL_TIPMOV = ? OR ALL_TIPMOV = ? OR  ALL_TIPMOV = ? ) AND ALL_FECHA_PRO >= ? AND ALL_FECHA_PRO <= ? AND ALL_FLAG_EXT <> 'E' AND ALL_CP = 'P'  "
  pub_cadena = "SELECT all_descto, ALL_IMPORTE_DOLL,ALL_NUMOPER,ALL_NUMOPER2, ALL_TIPMOV, ALL_MONEDA_CLI, ALL_NUMSER, ALL_NUMFAC, ALL_CP, ALL_IMPG2, ALL_IMPG1, ALL_CTAG2 , ALL_CTAG1, ALL_GASTOS, ALL_CODSUNAT, ALL_CODCLIE, ALL_FECHA_DIA, ALL_FECHA_SUNAT , ALL_IMPORTE_AMORT , ALL_IMPORTE ,ALL_BRUTO,ALL_IMPTO, ALL_NUMSER_C, ALL_NUMFAC_C, ALL_CODCIA FROM ALLOG, CLIENTES WHERE (ALL_RUC = CLI_RUC_ESPOSO) AND (ALL_CODCIA = CLI_CODCIA) AND (ALL_CODCIA = ? OR ALL_CODCIA = ? OR ALL_CODCIA = ? ) AND (ALL_TIPMOV = ? OR ALL_TIPMOV = ?  OR  ALL_TIPMOV = ? ) AND ALL_FECHA_PRO >= ? AND ALL_FECHA_PRO <= ? AND ALL_FLAG_EXT <> 'E' AND ALL_CP = 'P' AND CLI_CP = 'P' AND ALL_CODCLIE <> 0 and all_numfac_c<>0"
End If
If opcompra(1).Value Then
  pub_cadena = "SELECT all_descto, ALL_FECHA_PRO ,ALL_IMPORTE_DOLL,ALL_NUMOPER,ALL_NUMOPER2,ALL_TIPMOV, ALL_MONEDA_CLI, ALL_NUMSER, ALL_NUMFAC, ALL_CP, ALL_IMPG2, ALL_IMPG1, ALL_CTAG2 , ALL_CTAG1, ALL_GASTOS, ALL_CODSUNAT, ALL_CODCLIE, ALL_FECHA_DIA, ALL_FECHA_SUNAT , ALL_IMPORTE_AMORT , ALL_IMPORTE ,ALL_BRUTO,ALL_IMPTO, ALL_NUMSER_C, ALL_NUMFAC_C, ALL_CODCIA FROM ALLOG, CLIENTES WHERE (ALL_RUC = CLI_RUC_ESPOSO) AND (ALL_CP = CLI_CP) AND (ALL_CODCIA = CLI_CODCIA) AND (ALL_CODCIA = ? OR ALL_CODCIA = ? OR ALL_CODCIA = ?) AND (ALL_TIPMOV = ? OR ALL_TIPMOV = ?  OR  ALL_TIPMOV = ? ) AND ALL_FECHA_PRO >= ? AND ALL_FECHA_PRO <= ? AND ALL_FLAG_EXT <> 'E' AND ALL_CP = 'P' AND CLI_CP = 'P' AND CLI_CODCLIE <> 0 and all_numfac_c<>0"
  If val(txt_cli.Text) <> 0 Then
     pub_cadena = pub_cadena + " AND CLI_RUC_ESPOSO = '" & LOC_RUC & "'"
  End If
End If
If opcompra(2).Value Then
  'pub_cadena = "SELECT * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ? OR ALL_CODCIA = ? ) AND (ALL_TIPMOV = ? OR ALL_TIPMOV = ? OR  ALL_TIPMOV = ? ) AND ALL_FECHA_PRO >= ? AND ALL_FECHA_PRO <= ? AND ALL_FLAG_EXT <> 'E' AND ALL_CP = 'P'  "
  pub_cadena = "SELECT all_descto, ALL_IMPORTE_DOLL,ALL_NUMOPER,ALL_NUMOPER2,ALL_TIPMOV, ALL_MONEDA_CLI, ALL_NUMSER, ALL_NUMFAC, ALL_CP, ALL_IMPG2, ALL_IMPG1, ALL_CTAG2 , ALL_CTAG1, ALL_GASTOS, ALL_CODSUNAT, ALL_CODCLIE, ALL_FECHA_DIA, ALL_FECHA_SUNAT , ALL_IMPORTE_AMORT , ALL_IMPORTE ,ALL_BRUTO,ALL_IMPTO, ALL_NUMSER_C, ALL_NUMFAC_C, ALL_CODCIA FROM ALLOG, CLIENTES WHERE (ALL_RUC = CLI_RUC_ESPOSO) AND (ALL_CODCIA = CLI_CODCIA) AND (ALL_CODCIA = ? OR ALL_CODCIA = ? OR ALL_CODCIA = ? ) AND (ALL_TIPMOV = ? OR ALL_TIPMOV = ?  OR  ALL_TIPMOV = ? ) AND ALL_FECHA_PRO >= ? AND ALL_FECHA_PRO <= ? AND ALL_FLAG_EXT <> 'E' AND ALL_CP = 'P' AND CLI_CP = 'P' AND ALL_CODCLIE <> 0  and all_numfac_c<>0"
  If val(cta1.Text) <> 0 And val(cta2.Text) <> 0 And val(cta3.Text) <> 0 Then
     pub_cadena = pub_cadena + " AND (ALL_CTAG1 = '" & cta1.Text & "' OR ALL_CTAG1 = '" & cta2.Text & "' OR ALL_CTAG1 = '" & cta3.Text & "' OR  ALL_CTAG2 = '" & cta1.Text & "' OR ALL_CTAG2 = '" & cta2.Text & "' OR ALL_CTAG2 = '" & cta3.Text & "' )"
  ElseIf val(cta1.Text) <> 0 And val(cta2.Text) <> 0 Then
     pub_cadena = pub_cadena + " AND (ALL_CTAG1 = '" & cta1.Text & "' OR ALL_CTAG1 = '" & cta2.Text & "' OR  ALL_CTAG2 = '" & cta1.Text & "' OR ALL_CTAG2 = '" & cta2.Text & "' )"
  ElseIf val(cta1.Text) <> 0 Then
     pub_cadena = pub_cadena + " AND (ALL_CTAG1 = '" & cta1.Text & "' OR  ALL_CTAG2 = '" & cta1.Text & "' )"
  End If
End If
If val(codsunat.Text) <> 0 Then
 pub_cadena = pub_cadena + " AND ALL_CODSUNAT = " & codsunat.Text
End If
If Trim(moneda.Text) <> "T" Then
  If Trim(moneda.Text) = "S" Then
    pub_cadena = pub_cadena + " AND ALL_MONEDA_CLI = '" & Trim(moneda.Text) & "'"
  ElseIf Trim(moneda.Text) = "D" Then
    pub_cadena = pub_cadena + " AND ALL_MONEDA_CLI = '" & Trim(moneda.Text) & "'"
  End If
End If

If Trim(txtorden.Text) = "F" Then
  pub_cadena = pub_cadena + " ORDER BY ALL_FECHA_SUNAT,ALL_NUMSER_C,ALL_NUMFAC_C "
ElseIf Trim(txtorden.Text) = "D" Then
  pub_cadena = pub_cadena + " ORDER BY ALL_CODSUNAT,ALL_NUMSER_C,ALL_NUMFAC_C"
ElseIf Trim(txtorden.Text) = "R" Then
  'pub_cadena = "SELECT ALL_TIPMOV, ALL_MONEDA_CLI, ALL_NUMSER, ALL_NUMFAC, ALL_CP, ALL_IMPG2, ALL_IMPG1, ALL_CTAG2 , ALL_CTAG1, ALL_GASTOS, ALL_CODSUNAT, ALL_CODCLIE, ALL_FECHA_DIA, ALL_FECHA_SUNAT , ALL_IMPORTE_AMORT , ALL_IMPORTE ,ALL_BRUTO,ALL_IMPTO, ALL_NUMSER_C, ALL_NUMFAC_C, ALL_CODCIA FROM ALLOG, CLIENTES WHERE (ALL_CODCLIE = CLI_CODCLIE) AND (ALL_CODCIA = CLI_CODCIA) AND (ALL_CODCIA = ? OR ALL_CODCIA = ? OR ALL_CODCIA = ? ) AND (ALL_TIPMOV = ? OR ALL_TIPMOV = ?  OR  ALL_TIPMOV = ? ) AND ALL_FECHA_PRO >= ? AND ALL_FECHA_PRO <= ? AND ALL_FLAG_EXT <> 'E' AND ALL_CP = 'P' "
  If val(codsunat.Text) <> 0 Then
    pub_cadena = pub_cadena + " AND ALL_CODSUNAT = " & codsunat.Text
  End If
  pub_cadena = pub_cadena + " ORDER BY CLI_RUC_ESPOSO"
End If

Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = 0
PS_REP01(3) = 0
PS_REP01(4) = 0
PS_REP01(5) = 0
PS_REP01(6) = LK_FECHA_DIA
PS_REP01(7) = LK_FECHA_DIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
'Debug.Print pub_cadena
PS_REP01(2) = 0
If checompras.Value = 1 Then
 PS_REP01(3) = 20
 PS_REP01(4) = -1
 PS_REP01(5) = -1
ElseIf chenc.Value = 1 Then
 PS_REP01(3) = 97
 PS_REP01(4) = -1
 PS_REP01(5) = -1
Else
 PS_REP01(3) = 20
 PS_REP01(4) = 99
 PS_REP01(5) = 97
End If
PS_REP01(6) = wsFECHA1
PS_REP01(7) = wsFECHA2

PS_REP01(0) = ""
If Trim(par_llave!par_art_cias) = "" Then
PS_REP01(0) = LK_CODCIA
GoTo sigue
End If
WEMPRESA = ""
If LISCIA.Selected(0) Then
  LISCIA.ListIndex = 0
  PS_REP01(0) = Left(LISCIA.Text, 2)
  PSPAR_MULTI(0) = PS_REP01(0)
  par_multi.Requery
  WEMPRESA = "-" & Trim(par_multi!par_nombre_corto)
  PS_REP01(1) = ""
End If
'bloqado por mic para diro
'If LISCIA.Selected(1) Then
' LISCIA.ListIndex = 1
' PS_REP01(1) = Left(LISCIA.Text, 2)
' PSPAR_MULTI(0) = PS_REP01(1)
' par_multi.Requery
' WEMPRESA = WEMPRESA + " -" & Trim(par_multi!par_nombre_corto)
'End If
sigue:
If WEMPRESA = "" Then
  WEMPRESA = Trim(par_llave!PAR_NOMBRE)
Else
  WEMPRESA = Trim(GEN!GEN_NOMBRE) & " " & WEMPRESA
End If
DoEvents
RCRYSTAL.lblproceso.Visible = True
RCRYSTAL.ProgBar.Visible = True
RCRYSTAL.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
llave_rep01.Requery
If llave_rep01.EOF Then
  MsgBox "No Existe Movimientos", 48, Pub_Titulo
  GoTo CANCELA
End If


FILTRO_CTA(1) = "730001"
FILTRO_CTA(2) = "609001"

RCRYSTAL.lblproceso.Caption = "Procesando . . . "
DoEvents
RCRYSTAL.ProgBar.Visible = True
DoEvents
RCRYSTAL.ProgBar.Min = 0
RCRYSTAL.ProgBar.Value = 0
RCRYSTAL.ProgBar.Max = llave_rep01.RowCount
IMP_MONEDA = ""
f1 = 5
Lini = 6
wsigno = 1
Do Until llave_rep01.EOF
  RCRYSTAL.ProgBar.Value = RCRYSTAL.ProgBar.Value + 1
'  If llave_rep01!ALL_NUMFAC = 27 Then Stop
'  Print llave_rep01!ALL_NUMOPER2
'  If llave_rep01!ALL_TIPMOV = 13 Then Stop
'  If llave_rep01!ALL_FLAG_EXT = 13 Then Stop
  DoEvents
  wsigno = 1
  If llave_rep01!ALL_tipmov = 97 Then
     wsigno = -1
     If val(llave_rep01!ALL_NUMOPER) <> val(llave_rep01!ALL_numoper2) Then GoTo pasa
  End If
  WTC = 1
  IMP_MONEDA = llave_rep01!ALL_MONEDA_CLI
  If Trim(moneda.Text) = "S" Or Trim(moneda.Text) = "D" Or Trim(moneda.Text) = "A" Then
  Else
    If llave_rep01!ALL_MONEDA_CLI = "D" Then
      WTC = JALAR(llave_rep01!ALL_FECHA_SUNAT)
    End If
  End If
  
  If cheigv.Value = 1 Then
    WCHE_TOTAL = redondea((val(llave_rep01!ALL_IMPORTE_AMORT) * WTC)) - redondea(val(llave_rep01!ALL_GASTOS) * WTC)
    WCHE_IGV = redondea(WCHE_TOTAL - (WCHE_TOTAL / (1 + LK_IGV / 100)))
    If Abs(redondea(val(llave_rep01!ALL_IMPTO) * WTC) - WCHE_IGV) >= val(difigv.Text) Then
'    Stop
    Else
      GoTo pasa
    End If
  End If
  
  f1 = f1 + 1
  xl.Cells(f1, 1) = "'" & Format(llave_rep01!ALL_FECHA_SUNAT, "dd/mm")
  
  xl.Cells(f1, 2) = "'" & Format(llave_rep01!all_numser_c, "000")
  xl.Cells(f1, 3) = "'" & Format(llave_rep01!all_numfac_c, "0000000")
  xl.Cells(f1, 4) = "'" & Format(llave_rep01!ALL_CODSUNAT, "00")
 ' Print llave_rep01!ALL_FECHA_PRO
  
  pu_cp = llave_rep01!ALL_CP
  pu_codcia = llave_rep01!all_CODCIA
  SQ_OPER = 1
  pu_codclie = llave_rep01!ALL_CODCLIE
  LEER_CLI_LLAVE
  xl.Cells(f1, 5) = Trim(cli_llave!CLI_NOMBRE)
  xl.Cells(f1, 6) = Trim(cli_llave!cli_ruc_esposo)
  xl.Cells(f1, 7) = llave_rep01!ALL_MONEDA_CLI
'  If Val(llave_rep01!ALL_NUMFAC_C) = 2872 Then Stop
  If llave_rep01!ALL_tipmov = 97 Then
    xl.Cells(f1, 12) = (redondea((val(llave_rep01!ALL_IMPORTE_DOLL) * WTC)) - redondea(val(llave_rep01!ALL_GASTOS) * WTC)) * wsigno
  Else
   If llave_rep01!ALL_GASTOS <> 0 Then
     xl.Cells(f1, 12) = (redondea((val(llave_rep01!ALL_IMPORTE_AMORT) * WTC))) - redondea(val(llave_rep01!ALL_GASTOS) * WTC)
   Else
     xl.Cells(f1, 12) = (redondea((val(llave_rep01!ALL_IMPORTE_AMORT) * WTC)) - redondea(val(llave_rep01!ALL_GASTOS) * WTC)) * wsigno
   End If
  End If
  xl.Cells(f1, 9) = redondea(val(llave_rep01!ALL_GASTOS) * WTC) * wsigno
  xl.Cells(f1, 10) = Trim(cli_llave!CLI_CUENTA_CONTAB)
  xl.Cells(f1, 11) = redondea(val(llave_rep01!ALL_IMPTO) * WTC) * wsigno
'  xl.Application.Visible = True
  If llave_rep01!ALL_tipmov = 20 Then
    xl.Cells(f1, 8) = redondea(val(llave_rep01!ALL_BRUTO * WTC)) * wsigno - redondea(val(llave_rep01!ALL_DESCTO) * WTC) * wsigno ' agregue mic redondea(Val(llave_rep01!ALL_descto) * WTC) * wsigno
  Else
    xl.Cells(f1, 8) = 0
  End If
  IMP_CTA1 = redondea(Nulo_Valor0(llave_rep01!ALL_IMPG1) * WTC) * wsigno
  IMP_CTA2 = redondea(Nulo_Valor0(llave_rep01!ALL_IMPG2) * WTC) * wsigno
  wDescto = 0
  wFLETE = 0
  wCTAR = 0
  fca1 = ""
  fca2 = ""
  wCTAR = 0
  wCTARCTA = ""
  wCTAR2 = 0
  wCTARCTA2 = ""
  If llave_rep01!ALL_tipmov = 97 Then
   'If Trim(FILTRO_CTA(1)) <> "" Then
   '  If Trim(llave_rep01!ALL_CTAG1) = Trim(FILTRO_CTA(1)) And IMP_CTA1 <> 0 Then
    wDescto = redondea(val(llave_rep01!ALL_BRUTO * WTC)) * wsigno
    wDesctoCTA = FILTRO_CTA(1)
    fca1 = "A"
   '    End If
   '    If Trim(llave_rep01!ALL_CTAG2) = Trim(FILTRO_CTA(1)) And IMP_CTA2 <> 0 Then
   '       wDescto = IMP_CTA2
   '       wDesctoCTA = Trim(llave_rep01!ALL_CTAG2)
   ''       fca2 = "A"
   '    End If
  End If
  If Trim(FILTRO_CTA(2)) <> "" Then
     If Trim(llave_rep01!ALL_CTAG1) = Trim(FILTRO_CTA(2)) And IMP_CTA1 <> 0 Then
        wFLETE = IMP_CTA1
        wFLETECTA = Trim(llave_rep01!ALL_CTAG1)
        fca1 = "A"
     End If
     If Trim(llave_rep01!ALL_CTAG2) = Trim(FILTRO_CTA(2)) And IMP_CTA2 <> 0 Then
        wFLETE = IMP_CTA2
        wFLETECTA = Trim(llave_rep01!ALL_CTAG2)
        fca2 = "A"
     End If
  End If
  If fca1 <> "A" Then
     If Trim(llave_rep01!ALL_CTAG1) <> "" And IMP_CTA1 <> 0 Then
        wCTAR = IMP_CTA1
        wCTARCTA = Trim(llave_rep01!ALL_CTAG1)
        fca1 = "A"
     End If
  End If
  If fca2 <> "A" Then
     If Trim(llave_rep01!ALL_CTAG2) <> "" And IMP_CTA2 <> 0 Then
        If wCTAR = 0 Then
          wCTAR = IMP_CTA2
          wCTARCTA = Trim(llave_rep01!ALL_CTAG2)
          fca2 = "A"
        Else
          wCTAR2 = IMP_CTA2
          wCTARCTA2 = Trim(llave_rep01!ALL_CTAG2)
          fca2 = "A"
        End If
     End If
  End If
  If fca1 <> "A" Then
     If Trim(llave_rep01!ALL_CTAG1) <> "" And IMP_CTA1 <> 0 Then
        wCTAR2 = IMP_CTA1
        wCTARCTA2 = Trim(llave_rep01!ALL_CTAG1)
        fca1 = "A"
     End If
  End If
  If fca2 <> "A" Then
     If Trim(llave_rep01!ALL_CTAG2) <> "" And IMP_CTA2 <> 0 Then
        wCTAR2 = IMP_CTA2
        wCTARCTA2 = Trim(llave_rep01!ALL_CTAG2)
        fca2 = "A"
     End If
  End If
  
  xl.Cells(f1, 13) = val(wDescto)
  xl.Cells(f1, 14) = val(wFLETE)
  xl.Cells(f1, 15) = wCTARCTA
  xl.Cells(f1, 16) = val(wCTAR)
  'CHE_IMPORTE = Val(xl.Cells(f1, 11)) + Val(xl.Cells(f1, 12)) + Val(xl.Cells(f1, 13)) + Val(xl.Cells(f1, 14)) + Val(xl.Cells(f1, 16))
  If wCTAR2 <> 0 Then
    f1 = f1 + 1
    xl.Cells(f1, 15) = wCTARCTA2
    xl.Cells(f1, 16) = val(wCTAR2)
  '  CHE_IMPORTE = CHE_IMPORTE + Val(xl.Cells(f1, 16))
  '  CHE_IMPORTE = (Val(xl.Cells(f1 - 1, 8)) + Val(xl.Cells(f1 - 1, 9))) - CHE_IMPORTE
  '  If WTC <> 1 And Val(xl.Cells(f1 - 1, 11)) <> 0 Then xl.Cells(f1 - 1, 11) = xl.Cells(f1 - 1, 11) + CHE_IMPORTE
  Else
  '  CHE_IMPORTE = Val(xl.Cells(f1, 8)) - CHE_IMPORTE
  '  If WTC <> 1 And Val(xl.Cells(f1, 11)) <> 0 Then xl.Cells(f1, 11) = xl.Cells(f1, 11) + CHE_IMPORTE
  End If
  
  xl.Cells(f1, 17) = "'" & llave_rep01!ALL_NUMSER & "-" & llave_rep01!all_numfac   ' NRO, INTERNO
  xl.Cells(f1, 18) = WTC '"'" & Format(llave_rep01!all_CODCIA, "00")
pasa:
 llave_rep01.MoveNext
'  xl.Application.Visible = True
Loop
' xl.Application.Visible = True
Lfin = f1
  f1 = f1 + 2
  xl.Cells(f1, 1) = "Total Genral = "
  wran1 = "H" & 6
  wran2 = "H" & f1 - 1
  wranF = "H" & f1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "I" & 6
  wran2 = "I" & f1 - 1
  wranF = "I" & f1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "K" & 6
  wran2 = "K" & f1 - 1
  wranF = "K" & f1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "L" & 6
  wran2 = "L" & f1 - 1
  wranF = "L" & f1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "M" & 6
  wran2 = "M" & f1 - 1
  wranF = "M" & f1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "N" & 6
  wran2 = "N" & f1 - 1
  wranF = "N" & f1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "M" & 6
  wran2 = "M" & f1 - 1
  wranF = "M" & f1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  
  RCRYSTAL.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(5, 11) = "401001"
  xl.Cells(5, 8) = "AFECTO"
  xl.Cells(5, 13) = FILTRO_CTA(1)
  xl.Cells(5, 14) = FILTRO_CTA(2)

  xl.Cells(1, 1) = WEMPRESA '
  xl.Cells(2, 1) = Trim(retra_llave!TRA_DESCRIPCION)
  xl.Cells(3, 1) = "'" & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
 ' pub_mensaje = "Desea mostrar el Resumen de Asiento Contable... ?"
 ' Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
 ' If Pub_Respuesta = vbYes Then
  '  GoSub PASAR_CONT

  If chepasa.Value = 1 Then
    ASIENTO_MOVICONT xl, 1
  End If
  
  
  
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ""
  xl.APPLICATION.Visible = True
  DoEvents
  RCRYSTAL.lblproceso.Visible = False
  RCRYSTAL.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  RCRYSTAL.Pantalla.Enabled = True
  RCRYSTAL.Pantalla.Caption = "Por &Pantalla"
  RCRYSTAL.lblproceso.Visible = False
  Pantalla.Enabled = True
  cmdCerrar.Enabled = True

Exit Sub



CANCELA:
  RCRYSTAL.Pantalla.Enabled = True
  RCRYSTAL.Pantalla.Caption = "Por &Pantalla"
  RCRYSTAL.lblproceso.Visible = False
  RCRYSTAL.ProgBar.Visible = False
  Pantalla.Enabled = True
  cmdCerrar.Enabled = True
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
  
  xl.Workbooks.Open PUB_RUTA_OTRO & "RCOMPRA.xls", 0, True, 4, "", ""
Return

FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
 GoTo CANCELA
 Resume Next
Exit Sub

PASAR_CONT:
' definicion de variables
Dim ts_codcta As String
Dim ts_suma As Currency
DoEvents
lblproceso.Caption = "Generando Resumen de Asiento Contable... "
DoEvents
RCRYSTAL.ProgBar.Min = 0
RCRYSTAL.ProgBar.Max = 5
RCRYSTAL.ProgBar.Value = 0
'---------------------------------------------
'*** Orden para toda la 42 o 46.. al Haber***
'---------------------------------------------
RCRYSTAL.ProgBar.Value = RCRYSTAL.ProgBar.Value + 1
wranF = "A" & Lini & ":S" & Lfin
xl.Sheets(1).Activate
  
'xl.Application.Visible = True
xl.APPLICATION.Worksheets(1).Range(wranF).Sort Key1:=xl.APPLICATION.Worksheets(1).Range("A1")
xl.APPLICATION.Worksheets(1).Range(wranF).Sort Key1:=xl.APPLICATION.Worksheets(1).Range("J1") ' , Key2:=xl.Application.Worksheets("Hoja1").Range("F1"), Key3:=xl.Application.Worksheets("Hoja1").Range("G1")
f1 = 4
fila = Lini
ts_codcta = Trim(Format(xl.Cells(fila, 10), "##########"))
ts_suma = 0
For fila = Lini To Lfin
If Trim(xl.Cells(fila, 5)) = "" Then GoTo cont_p
  If Trim(ts_codcta) <> Trim(Format(xl.Cells(fila, 10), "##########")) Then
    xl.Sheets(2).Activate
    f1 = f1 + 1
    xl.Cells(f1, 1) = ts_codcta
    xl.Cells(f1, 2) = JALA_CTA(ts_codcta)
    xl.Cells(f1, 3) = 0 ' debe
    xl.Cells(f1, 4) = ts_suma  ' haber
    xl.Cells(f1, 5) = "H"
    xl.Sheets(1).Activate
    ts_codcta = Trim(Format(xl.Cells(fila, 10), "##########"))
    ts_suma = 0
  End If
  ts_suma = ts_suma + (val(Format(xl.Cells(fila, 12), "0.000")) + val(Format(xl.Cells(fila, 9), "0.000")))
cont_p:
Next fila
xl.Sheets(2).Activate
xl.Cells(1, 1) = WEMPRESA '
xl.Cells(2, 1) = "PERIODO: '" & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")

f1 = f1 + 1
xl.Cells(f1, 1) = ts_codcta
xl.Cells(f1, 2) = JALA_CTA(ts_codcta)
xl.Cells(f1, 3) = 0
xl.Cells(f1, 4) = ts_suma
xl.Cells(f1, 5) = "H"
ts_suma = 0
RCRYSTAL.ProgBar.Value = RCRYSTAL.ProgBar.Value + 1
'---------------------------------------------
'*** coloca el total para el igv y  Mercaderia.. al Debe***
'---------------------------------------------
' coloca la cta  de mercaedria 60..
xl.Sheets(1).Activate
ts_codcta = Trim(Format(xl.Cells(5, 8), "##########"))
ts_suma = val(Format(xl.Cells(Lfin + 2, 8), "0.000"))
f1 = f1 + 1
xl.Sheets(2).Activate
xl.Cells(f1, 1) = ts_codcta
xl.Cells(f1, 2) = JALA_CTA(ts_codcta)
xl.Cells(f1, 3) = ts_suma ' debe
xl.Cells(f1, 4) = 0 ' haber
xl.Cells(f1, 5) = "D"
xl.Sheets(1).Activate
' coloca la cta  del IGV 40..
ts_codcta = Trim(Format(xl.Cells(5, 11), "##########"))
ts_suma = val(Format(xl.Cells(Lfin + 2, 11), "0.000"))
xl.Sheets(2).Activate
f1 = f1 + 1
xl.Cells(f1, 1) = ts_codcta
xl.Cells(f1, 2) = JALA_CTA(ts_codcta)
xl.Cells(f1, 3) = ts_suma ' debe
xl.Cells(f1, 4) = 0 ' haber
xl.Cells(f1, 5) = "D"
xl.Sheets(1).Activate
' coloca la cta  de Nota de Credito  73..
ts_codcta = Trim(Format(xl.Cells(5, 13), "##########"))
ts_suma = val(Format(xl.Cells(Lfin + 2, 13), "0.000"))
xl.Sheets(2).Activate
f1 = f1 + 1
xl.Cells(f1, 1) = ts_codcta
xl.Cells(f1, 2) = JALA_CTA(ts_codcta)
xl.Cells(f1, 3) = ts_suma ' debe
xl.Cells(f1, 4) = 0 ' haber
xl.Cells(f1, 5) = "D"
xl.Sheets(1).Activate
' coloca la cta  de Fletes  609..
ts_codcta = Trim(Format(xl.Cells(5, 14), "##########")) ' Trim(xl.Cells(5, 14))
ts_suma = val(Format(xl.Cells(Lfin + 2, 14), "0.000"))
xl.Sheets(2).Activate
f1 = f1 + 1
xl.Cells(f1, 1) = ts_codcta
xl.Cells(f1, 2) = JALA_CTA(ts_codcta)
xl.Cells(f1, 3) = ts_suma ' debe
xl.Cells(f1, 4) = 0 ' haber
xl.Cells(f1, 5) = "D"
xl.Sheets(1).Activate
ts_suma = 0
RCRYSTAL.ProgBar.Value = RCRYSTAL.ProgBar.Value + 1
'---------------------------------------------
'*** Orden para toda la 6... al Debe***
'---------------------------------------------
wranF = "A" & Lini & ":S" & Lfin
xl.Sheets(1).Activate
'xl.Application.Visible = True
xl.APPLICATION.Worksheets(1).Range(wranF).Sort Key1:=xl.APPLICATION.Worksheets(1).Range("A1")
xl.APPLICATION.Worksheets(1).Range(wranF).Sort Key1:=xl.APPLICATION.Worksheets(1).Range("O1")
fila = Lini
ts_codcta = Trim(Format(xl.Cells(fila, 15), "##########"))  ' Trim(xl.Cells(fila, 15))
ts_suma = 0
For fila = Lini To Lfin
If Trim(xl.Cells(fila, 5)) = "" And Trim(xl.Cells(fila, 15)) = "" Then GoTo cont_p1
  If Trim(ts_codcta) <> Trim(Format(xl.Cells(fila, 15), "##########")) Then
    xl.Sheets(2).Activate
    f1 = f1 + 1
    xl.Cells(f1, 1) = ts_codcta
    xl.Cells(f1, 2) = JALA_CTA(ts_codcta)
    xl.Cells(f1, 3) = ts_suma  ' debe
    xl.Cells(f1, 4) = 0  ' haber
    xl.Cells(f1, 5) = "D"
    xl.Sheets(1).Activate
    ts_codcta = Trim(Format(xl.Cells(fila, 15), "##########"))  'Trim(xl.Cells(fila, 15))
    ts_suma = 0
  End If
  ts_suma = ts_suma + (val(Format(xl.Cells(fila, 16), "0.000")))
cont_p1:
Next fila
xl.Sheets(2).Activate
If Trim(ts_codcta) <> "" Then
 f1 = f1 + 1
 xl.Cells(f1, 1) = ts_codcta
 xl.Cells(f1, 2) = JALA_CTA(ts_codcta)
 xl.Cells(f1, 3) = ts_suma
 xl.Cells(f1, 4) = 0
 xl.Cells(f1, 5) = "D"
 ts_suma = 0
End If

' TOTLES Y ORDEN DE ASIENTO
IMP_CTA1 = 0
RCRYSTAL.ProgBar.Value = RCRYSTAL.ProgBar.Value + 1
wran1 = "C" & 5
wran2 = "C" & f1
wranF = "C" & f1 + 1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
IMP_CTA1 = IMP_CTA1 + val(xl.Range(wranF))
wran1 = "D" & 5
wran2 = "D" & f1
wranF = "D" & f1 + 1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
IMP_CTA1 = IMP_CTA1 - val(xl.Range(wranF))

wranF = "C" & f1 + 3
xl.Range(wranF) = IMP_CTA1
wranF = "B" & f1 + 3
xl.Range(wranF) = "Diferencia:"


wranF = "A" & 5 & ":S" & f1
xl.APPLICATION.Worksheets(2).Range(wranF).Sort Key1:=xl.APPLICATION.Worksheets(2).Range("A1")
xl.APPLICATION.Worksheets(2).Range(wranF).Sort Key1:=xl.APPLICATION.Worksheets(2).Range("E1")
'RCRYSTAL.ProgBar.Value = RCRYSTAL.ProgBar.Value + 1
    
Return

End Sub

Private Sub txtorden_Change()
txtorden.Text = UCase(txtorden.Text)
End Sub

Private Sub txtorden_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If UCase(txtorden.Text) = "F" Or UCase(txtorden.Text) = "D" Or UCase(txtorden.Text) = "R" Then
Else
  MsgBox "NO PROCEDE..", 48, Pub_Titulo
  Azul txtorden, txtorden
  KeyAscii = 0
  Exit Sub
End If
If Pantalla.Enabled Then Pantalla.SetFocus

End Sub

Public Sub REG_BANCOS()
'  ******** Centralizador de Bancos. *********
On Error GoTo FINTODO
Dim ts_suma As Currency
Dim ts_DH As String
Dim ts_codcta As String
Dim Lini As Integer
Dim Lfin As Integer
Dim qver_onlyCont As Integer
Dim WP_NOTA_PROV  As Currency
Dim WP_DOC_PROV As Currency
Dim DES_CTA_FAVOR As String
Dim DES_CTA_CONTRA As String
Dim DES_COD_FAVOR As String
Dim DES_COD_CONTRA As String
Dim RB_CTA As String
Dim RB_DESCRIPCION_DIF As String
Dim RB_CTACONT_DIF As String
Dim IMP_CONTAB As Currency
Dim IMPORTE_DIF As Currency
Dim WP_NOTAC  As Currency
Dim WP_PROV  As Currency
Dim W_IMPORTE As Currency
Dim wpasa_rep As Integer
Dim wsigno As Currency
Dim WEMPRESA As String
Dim WCHE_TOTAL As Currency
Dim WCHE_IGV As Currency
Dim xcuenta  As Integer
Dim wCTAR2 As Currency
Dim wCTARCTA2  As String * 12
Dim fca1 As String * 1
Dim fca2 As String * 1
Dim IMP_CTA1 As Currency
Dim IMP_CTA2 As Currency
Dim WTC As Currency
Dim wOTRO As Currency
Dim wOTROCTA As String
Dim wFLETE As Currency
Dim wFLETECTA As String
Dim wCTAR As Currency
Dim wCTARCTA As String
Dim wDescto  As Currency
Dim wDesctoCTA As String
Dim LETRAS(100) As String * 2
Dim FILTRO_CTA(3) As String
Dim wsFECHA1
Dim wsFECHA2
Dim wcta1 As String
Dim wIMPORTE1 As Currency
Dim IMP_MONEDA As String * 1
Dim ts_sumaH  As Currency
Dim ts_sumaD As Currency

'NUEVAS***********
Dim xlR  As Object
Dim CTACONT As String
Dim RB_CTACONT As String
Dim RB_COMPRO As String
Dim RB_DESCRIPCION As String
Dim RB_SECUENCIA As Currency
Dim RB_TIPO  As String
Dim RB_NUMSER_C As Integer
Dim RB_NUMFAC_C As Currency
Dim RB_CONCEPTO As String
Dim RB_CARGO As Currency
Dim RB_ABANO As Currency
Dim RB_NOMCORTO As String

Dim fin_filas As Integer
Dim imp_cargo As Currency
Dim imp_abono As Currency
Dim tot_cargo As Currency
Dim tot_abono As Currency



Pantalla.Enabled = False
cmdCerrar.Enabled = False
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
 Azul2 txtCampo2, txtCampo2
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If chepasa.Value = 1 Then
  pub_mensaje = "<Advertencia> El pase de la información es por cada Compañia. Continuar...?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then
      Azul2 txtCampo1, txtCampo1
      GoTo CANCELA
  End If
  qver_onlyCont = CHE_BLOQ_MES(3)
  If qver_onlyCont = 1 Then
    MsgBox "Periodo Activo esta Cerrado. No procede.", 48, Pub_Titulo
    chepasa.Value = 0
    GoTo CANCELA
  End If
  If (cop_llave!cop_fecha_proceso = CDate(wsFECHA1)) And (cop_llave!cop_fecha_proceso2 = CDate(wsFECHA2)) Then
  Else
    MsgBox "Usted. a marcado la opción: Pasar la Información al Periodo Contable. " & Chr(13) & "Las Fechas ingresadas son distintas a la del Periodo Contable . Verificar...", 48, Pub_Titulo
    Azul2 txtCampo1, txtCampo1
    GoTo CANCELA
  End If
  If qver_onlyCont = 9 Then ' hay Información en OnlyCont. Confirmar.
    pub_mensaje = "Usted. a marcado la opción: Pasar la Información al Periodo Contable. " & Chr(13) & Chr(13) & "Existe Información en este Periodo Contable , el Sistema Reemplazazá la Información.  " & Chr(13) & Chr(13) & "<Desea Continuar de todas maneras>...?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
      Azul2 txtCampo1, txtCampo1
      GoTo CANCELA
    End If
    
  End If
  
End If

'If CDate(wsFECHA1) <> cop_llave!cop_fecha_proceso Then cheasiento.Value = 0
'If CDate(wsFECHA2) <> cop_llave!COP_FECHA_PROCESO2 Then cheasiento.Value = 0
GoSub WEXCEL
pub_cadena = ""
'xl.Application.Visible = True
xcuenta = 0
SQ_OPER = 1
PUB_CUENTA = Nulo_Valors(cop_llave!COP_CTA_DIF_TC_CONTRA)
DES_COD_CONTRA = PUB_CUENTA
LEER_COM_LLAVE
If com_llave.EOF Then
 DES_CTA_CONTRA = "Cta. No Definida..."
Else
 DES_CTA_CONTRA = com_llave!com_descripcion
End If
PUB_CUENTA = Nulo_Valors(cop_llave!COP_CTA_DIF_TC_FAVOR)
DES_COD_FAVOR = PUB_CUENTA
LEER_COM_LLAVE
If com_llave.EOF Then
 DES_CTA_FAVOR = "Cta. No Definida..."
Else
 DES_CTA_FAVOR = com_llave!com_descripcion
End If


Pantalla.Enabled = False
cmdCerrar.Enabled = False
DoEvents
RCRYSTAL.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
pub_cadena = ""
wpasa_rep = 0
f1 = 0
pub_cadena = "SELECT * FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_NUMSER_C = ? AND ALL_NUMFAC_C = ?  AND ALL_CODCLIE = ? AND ALL_CP = ? AND ALL_SIGNO_CAR = 1 AND ALL_FLAG_EXT <> 'E'"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
PS_REP02(1) = 0
PS_REP02(2) = 0
PS_REP02(3) = 0
PS_REP02(4) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT CAA_SALDO_CAR, CAA_IMPORTE, CAA_FECHA_COBRO FROM CARACU WHERE CAA_CODCIA = ? AND CAA_CP = ? AND CAA_CODCLIE = ? AND CAA_SERDOC = ? AND CAA_NUMDOC = ?  AND CAA_NOTA = 'C'  AND CAA_ESTADO <> 'E' "
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
PS_REP03(0) = 0
PS_REP03(1) = 0
PS_REP03(2) = 0
PS_REP03(3) = 0
PS_REP03(4) = 0
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


OTRO_PASE:
wpasa_rep = wpasa_rep + 1
If wpasa_rep = 1 Then
    If val(txt_key.Text) <> 0 Then
      pub_cadena = "SELECT ALL_SECUENCIA, ALL_TIPDOC, ALL_FBG, ALL_NUMGUIA, ALL_TIPO_CAMBIO, ALL_CTAG1, ALL_MONEDA_CCM, ALL_FECHA_CAN, ALL_SIGNO_CCM, ALL_CODBAN, ALL_CODTRA, all_signo_ccm , all_concepto, ALL_CHENUM, CLI_CUENTA_CONTAB,CLI_NOMBRE, ALL_IMPORTE_DOLL,ALL_MONEDA_CLI, ALL_NUMSER, ALL_NUMFAC, ALL_GASTOS, ALL_CODSUNAT, ALL_CODCLIE, ALL_FECHA_DIA, ALL_FECHA_SUNAT, ALL_IMPORTE_AMORT , ALL_IMPORTE ,ALL_BRUTO,ALL_IMPTO, ALL_NUMSER_C, ALL_NUMFAC_C, ALL_CODCIA FROM ALLOG, CLIENTES WHERE (ALL_RUC = CLI_RUC_ESPOSO) AND (ALL_CODCIA = CLI_CODCIA) AND (ALL_CP = CLI_CP) AND (ALL_CODCIA = ? OR ALL_CODCIA = ? OR ALL_CODCIA = ? ) AND (ALL_CODTRA = 2748 OR ALL_CODTRA = 2735 OR ALL_CODTRA = 5714 OR ALL_CODTRA = 2738 ) AND (ALL_FECHA_CAN >= ? AND ALL_FECHA_CAN <= ?) AND (ALL_SIGNO_CCM <> 0 AND ALL_FLAG_EXT <> 'E' AND ALL_CODBAN = " & Trim(txt_key.Text) & ") ORDER BY CLI_CUENTA_CONTAB, ALL_CTAG1 "
    Else
      pub_cadena = "SELECT ALL_SECUENCIA, ALL_TIPDOC, ALL_FBG, ALL_NUMGUIA, ALL_TIPO_CAMBIO, ALL_CTAG1, ALL_MONEDA_CCM, ALL_FECHA_CAN, ALL_SIGNO_CCM, ALL_CODBAN, ALL_CODTRA, all_signo_ccm , all_concepto, ALL_CHENUM, CLI_CUENTA_CONTAB,CLI_NOMBRE, ALL_IMPORTE_DOLL,ALL_MONEDA_CLI, ALL_NUMSER, ALL_NUMFAC, ALL_GASTOS, ALL_CODSUNAT, ALL_CODCLIE, ALL_FECHA_DIA, ALL_FECHA_SUNAT, ALL_IMPORTE_AMORT , ALL_IMPORTE ,ALL_BRUTO,ALL_IMPTO, ALL_NUMSER_C, ALL_NUMFAC_C, ALL_CODCIA FROM ALLOG, CLIENTES WHERE (ALL_RUC = CLI_RUC_ESPOSO) AND (ALL_CODCIA = CLI_CODCIA) AND (ALL_CP = CLI_CP) AND (ALL_CODCIA = ? OR ALL_CODCIA = ? OR ALL_CODCIA = ?) AND (ALL_FECHA_CAN >= ? AND ALL_FECHA_CAN <= ?) AND (ALL_CODTRA = 2748 OR ALL_CODTRA = 2735 OR ALL_CODTRA = 5714 OR ALL_CODTRA = 2738 ) AND ALL_SIGNO_CCM <> 0 AND ALL_FLAG_EXT <> 'E'  ORDER BY CLI_CUENTA_CONTAB, ALL_CTAG1 "
    End If
Else
    If val(txt_key.Text) <> 0 Then
      pub_cadena = "SELECT ALL_SECUENCIA, ALL_TIPDOC, ALL_FBG, ALL_NUMGUIA, ALL_TIPO_CAMBIO, ALL_CTAG1, ALL_MONEDA_CCM, ALL_FECHA_CAN, ALL_SIGNO_CCM, ALL_CODBAN, ALL_CODTRA, all_signo_ccm , all_concepto, ALL_CHENUM, CLI_CUENTA_CONTAB,CLI_NOMBRE, ALL_IMPORTE_DOLL,ALL_MONEDA_CLI, ALL_NUMSER, ALL_NUMFAC, ALL_GASTOS, ALL_CODSUNAT, ALL_CODCLIE, ALL_FECHA_DIA, ALL_FECHA_SUNAT, ALL_IMPORTE_AMORT , ALL_IMPORTE ,ALL_BRUTO,ALL_IMPTO, ALL_NUMSER_C, ALL_NUMFAC_C, ALL_CODCIA FROM ALLOG, CLIENTES WHERE (ALL_RUC = CLI_RUC_ESPOSO) AND (ALL_CODCIA = CLI_CODCIA) AND (ALL_CP = CLI_CP) AND (ALL_CODCIA = ? OR ALL_CODCIA = ? OR ALL_CODCIA = ? ) AND (ALL_CODTRA = 2720 OR ALL_CODTRA = 5318) AND (ALL_FECHA_CAN >= ? AND ALL_FECHA_CAN <= ?) AND (ALL_SIGNO_CCM <> 0 AND ALL_FLAG_EXT <> 'E' AND ALL_CODBAN = " & Trim(txt_key.Text) & ") ORDER BY CLI_CUENTA_CONTAB, ALL_CTAG1 "
    Else
      pub_cadena = "SELECT ALL_SECUENCIA, ALL_TIPDOC, ALL_FBG, ALL_NUMGUIA, ALL_TIPO_CAMBIO, ALL_CTAG1, ALL_MONEDA_CCM, ALL_FECHA_CAN, ALL_SIGNO_CCM, ALL_CODBAN, ALL_CODTRA, all_signo_ccm , all_concepto, ALL_CHENUM, CLI_CUENTA_CONTAB,CLI_NOMBRE, ALL_IMPORTE_DOLL,ALL_MONEDA_CLI, ALL_NUMSER, ALL_NUMFAC, ALL_GASTOS, ALL_CODSUNAT, ALL_CODCLIE, ALL_FECHA_DIA, ALL_FECHA_SUNAT, ALL_IMPORTE_AMORT , ALL_IMPORTE ,ALL_BRUTO,ALL_IMPTO, ALL_NUMSER_C, ALL_NUMFAC_C, ALL_CODCIA FROM ALLOG, CLIENTES WHERE (ALL_RUC = CLI_RUC_ESPOSO) AND (ALL_CODCIA = CLI_CODCIA) AND (ALL_CP = CLI_CP) AND (ALL_CODCIA = ? OR ALL_CODCIA = ? OR ALL_CODCIA = ?) AND (ALL_FECHA_CAN >= ? AND ALL_FECHA_CAN <= ?) AND (ALL_CODTRA = 2720 OR ALL_CODTRA = 5318) AND ALL_SIGNO_CCM <> 0 AND ALL_FLAG_EXT <> 'E'  ORDER BY CLI_CUENTA_CONTAB, ALL_CTAG1 "
    End If
End If
    
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = 0
PS_REP01(3) = LK_FECHA_DIA
PS_REP01(4) = LK_FECHA_DIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = ""
PS_REP01(1) = ""
PS_REP01(2) = ""
PS_REP01(3) = CDate(wsFECHA1)
PS_REP01(4) = CDate(wsFECHA2)
PS_REP01(0) = ""
If Trim(par_llave!par_art_cias) = "" Then
PS_REP01(0) = LK_CODCIA
GoTo sigue
End If
WEMPRESA = ""
If LISCIA.Selected(0) Then
  LISCIA.ListIndex = 0
  PS_REP01(0) = Left(LISCIA.Text, 2)
  PSPAR_MULTI(0) = PS_REP01(0)
  par_multi.Requery
  WEMPRESA = "-" & Trim(par_multi!par_nombre_corto)
  PS_REP01(1) = ""
End If

If LISCIA.Selected(1) Then
 LISCIA.ListIndex = 1
 PS_REP01(1) = Left(LISCIA.Text, 2)
 PSPAR_MULTI(0) = PS_REP01(1)
 par_multi.Requery
 WEMPRESA = WEMPRESA + " -" & Trim(par_multi!par_nombre_corto)
End If
sigue:

DoEvents
RCRYSTAL.lblproceso.Visible = True
RCRYSTAL.ProgBar.Visible = True
RCRYSTAL.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
llave_rep01.Requery
If llave_rep01.EOF Then

   If wpasa_rep > 2 Then
   Else
    GoTo OTRO_PASE
   End If
  ' If wpasa_rep = 3 Then
   If f1 <> 0 Then
  '   GoTo OTRO_PASE
   Else
   
  ' End If
   MsgBox "No Existe Movimientos", 48, Pub_Titulo
   GoTo CANCELA
   End If
  
End If

RCRYSTAL.lblproceso.Caption = "Procesando Información. . . "
DoEvents
RCRYSTAL.ProgBar.Visible = True
DoEvents
RCRYSTAL.ProgBar.Min = 0
RCRYSTAL.ProgBar.Value = 0
If Not llave_rep01.EOF Then RCRYSTAL.ProgBar.Max = llave_rep01.RowCount
IMP_MONEDA = ""

wsigno = 1
CTACONT = -1 'llave_rep01!CLI_CUENTA_CONTAB
Do Until llave_rep01.EOF
 'If llave_rep01!all_CODCLIE = 3838 Then Stop
  RCRYSTAL.ProgBar.Value = RCRYSTAL.ProgBar.Value + 1
  DoEvents
  RB_COMPRO = Trim(llave_rep01!all_chenum)
  RB_DESCRIPCION = llave_rep01!CLI_NOMBRE
  RB_SECUENCIA = llave_rep01!all_numfac
  RB_TIPO = llave_rep01!all_CODCIA
  RB_NUMSER_C = llave_rep01!all_numser_c
  RB_NUMFAC_C = llave_rep01!all_numfac_c
  RB_CONCEPTO = llave_rep01!all_concepto
'  RB_CONCEPTO = llave_rep01!all_autocon
  SQ_OPER = 1
  pu_codcia = LK_CODCIA
  PUB_CODBAN = llave_rep01!all_codban
  LEER_CCM_LLAVE
  RB_NOMCORTO = ""
  RB_CTA = ""
  If Not ccm_llave.EOF Then
    RB_NOMCORTO = Trim(Nulo_Valors(ccm_llave!ccm_nomcorto))
    RB_CTA = Trim(Nulo_Valors(ccm_llave!ccm_cuenta_contab2))
  End If
  RB_CARGO = 0
  RB_ABANO = 0
  If llave_rep01!ALL_CODTRA = 5714 Then
    RB_CTACONT = Trim(llave_rep01!ALL_CTAG1)
    SQ_OPER = 1
    PUB_CUENTA = RB_CTACONT
    LEER_COM_LLAVE
    If com_llave.EOF Then
      RB_DESCRIPCION = "Cta. No Definida..."
    Else
      RB_DESCRIPCION = com_llave!com_descripcion
    End If
  ElseIf llave_rep01!ALL_CODTRA = 2720 Then
    SQ_OPER = 1
    pu_cp = "C"
    pu_codclie = llave_rep01!ALL_CODCLIE
    pu_codcia = llave_rep01!all_CODCIA
    LEER_CLI_LLAVE
    If Not cli_llave.EOF Then
      RB_CTACONT = Trim(cli_llave!CLI_CUENTA_CONTAB)
      If val(RB_NUMFAC_C) <> 0 Then
       RB_CONCEPTO = Trim(cli_llave!CLI_NOMBRE)
      End If
    End If
    SQ_OPER = 1
    PUB_CUENTA = RB_CTACONT
    LEER_COM_LLAVE
    If com_llave.EOF Then
      RB_DESCRIPCION = "Cta. No Definida..."
    Else
      RB_DESCRIPCION = com_llave!com_descripcion
    End If
  ElseIf llave_rep01!ALL_CODTRA = 5318 Then
    RB_CTACONT = Trim(Nulo_Valors(ccm_llave!CCM_CUENTA_CONTAB))
    SQ_OPER = 1
    PUB_CUENTA = RB_CTACONT
    LEER_COM_LLAVE
    If com_llave.EOF Then
      RB_DESCRIPCION = "Cta. No Definida..."
    Else
      RB_DESCRIPCION = com_llave!com_descripcion
    End If
  ElseIf llave_rep01!ALL_CODTRA = 2748 Or llave_rep01!ALL_CODTRA = 2735 Or llave_rep01!ALL_CODTRA = 2738 Then
    SQ_OPER = 1
    pu_cp = "P"
    pu_codclie = llave_rep01!ALL_CODCLIE
    pu_codcia = llave_rep01!all_CODCIA
    LEER_CLI_LLAVE
    If Not cli_llave.EOF Then
      RB_CTACONT = Trim(cli_llave!CLI_CUENTA_CONTAB)
      If Trim(RB_CONCEPTO) = "" Then
       RB_CONCEPTO = Trim(cli_llave!CLI_NOMBRE)
      End If
    End If
    SQ_OPER = 1
    PUB_CUENTA = RB_CTACONT
    LEER_COM_LLAVE
    If com_llave.EOF Then
      RB_DESCRIPCION = "Cta. No Definida..."
    Else
      RB_DESCRIPCION = com_llave!com_descripcion
    End If
    'RB_CTACONT = Trim(cli_llave!CLI_CUENTA_CONTAB) ' Trim(llave_rep01!CLI_CUENTA_CONTAB)
  End If
  
  If RB_CTACONT = "" Then GoTo dale1
  If Trim(cta1.Text) <> "" And Trim(cta2.Text) <> "" And Trim(cta3.Text) <> "" Then
    If (Trim(cta1.Text) = RB_CTACONT) Or (Trim(cta2.Text) = RB_CTACONT) Or (Trim(cta3.Text) = RB_CTACONT) Then
    Else
     GoTo pasa
    End If
  ElseIf Trim(cta1.Text) <> "" And Trim(cta2.Text) <> "" Then
    If (Trim(cta1.Text) = RB_CTACONT) Or (Trim(cta2.Text) = RB_CTACONT) Then
    Else
     GoTo pasa
    End If
  ElseIf Trim(cta1.Text) <> "" Then
    If (Trim(cta1.Text) = RB_CTACONT) Then
    Else
     GoTo pasa
    End If
  End If
dale1:
  f1 = f1 + 1
  W_IMPORTE = 0
  ' PRUEBAS VER
     If val(llave_rep01!all_chenum) = 80000411 Then Stop
  ' *******
     If llave_rep01!ALL_moneda_ccm = "D" Then
           If llave_rep01!ALL_CODTRA = 2748 Then
               W_IMPORTE = redondea(val(llave_rep01!ALL_IMPORTE) * JALAR_TC(llave_rep01!ALL_FECHA_SUNAT, 3))
               'VERIFICA SI EXISTE DIF. T.C.
               IMP_CONTAB = redondea(val(llave_rep01!ALL_IMPORTE) * JALAR_TC(llave_rep01!ALL_FECHA_CAN, llave_rep01!ALL_SIGNO_CCM))
               IMPORTE_DIF = W_IMPORTE - IMP_CONTAB
               If IMPORTE_DIF < 0 Then
                    RB_CTACONT_DIF = DES_COD_CONTRA
                    RB_CARGO = Abs(IMPORTE_DIF)
                    RB_DESCRIPCION_DIF = DES_CTA_CONTRA
                    GoSub ADI_DIF
               ElseIf IMPORTE_DIF > 0 Then
                    RB_CTACONT_DIF = DES_COD_FAVOR
                    RB_ABANO = Abs(IMPORTE_DIF)
                    RB_DESCRIPCION_DIF = DES_CTA_FAVOR
                    GoSub ADI_DIF
               End If
           ElseIf llave_rep01!ALL_CODTRA = 5714 And llave_rep01!ALL_SIGNO_CCM = 1 And llave_rep01!ALL_IMPORTE_AMORT <> 0 Then
               W_IMPORTE = redondea(val(llave_rep01!ALL_IMPORTE_AMORT))
               'VERIFICA SI EXISTE DIF. T.C.
               IMP_CONTAB = redondea(val(llave_rep01!ALL_IMPORTE) * JALAR_TC(llave_rep01!ALL_FECHA_CAN, llave_rep01!ALL_SIGNO_CCM))
               IMPORTE_DIF = W_IMPORTE - IMP_CONTAB
               If llave_rep01!ALL_SIGNO_CCM = -1 Then
                  If IMPORTE_DIF < 0 Then
                     RB_CTACONT_DIF = DES_COD_CONTRA
                     RB_CARGO = Abs(IMPORTE_DIF)
                     RB_DESCRIPCION_DIF = DES_CTA_CONTRA
                    GoSub ADI_DIF
                  ElseIf IMPORTE_DIF > 0 Then
                    RB_CTACONT_DIF = DES_COD_FAVOR
                    RB_ABANO = Abs(IMPORTE_DIF)
                    RB_DESCRIPCION_DIF = DES_CTA_FAVOR
                    GoSub ADI_DIF
                  End If
               Else
                  If IMPORTE_DIF > 0 Then
                     RB_CTACONT_DIF = DES_COD_CONTRA
                     RB_CARGO = Abs(IMPORTE_DIF)
                     RB_DESCRIPCION_DIF = DES_CTA_CONTRA
                    GoSub ADI_DIF
                  ElseIf IMPORTE_DIF < 0 Then
                    RB_CTACONT_DIF = DES_COD_FAVOR
                    RB_ABANO = Abs(IMPORTE_DIF)
                    RB_DESCRIPCION_DIF = DES_CTA_FAVOR
                    GoSub ADI_DIF
                  End If
               End If
           ElseIf llave_rep01!ALL_CODTRA = 5318 Then
              If val(llave_rep01!ALL_IMPORTE) <> val(llave_rep01!ALL_IMPORTE_AMORT) Then
                W_IMPORTE = redondea(val(llave_rep01!ALL_IMPORTE))
              Else
                W_IMPORTE = redondea(val(llave_rep01!ALL_IMPORTE) * JALAR_TC(llave_rep01!ALL_FECHA_CAN, -1))
              End If
              WP_DOC_PROV = redondea(val(llave_rep01!ALL_IMPORTE) * JALAR_TC(llave_rep01!ALL_FECHA_CAN, llave_rep01!ALL_SIGNO_CCM))
              IMP_CONTAB = WP_DOC_PROV
              If val(llave_rep01!all_SECUENCIA) = 1 Then
                IMPORTE_DIF = IMP_CONTAB - W_IMPORTE
              Else
                IMPORTE_DIF = W_IMPORTE - IMP_CONTAB
              End If
              If IMPORTE_DIF < 0 Then
                  RB_CTACONT_DIF = DES_COD_CONTRA
                  RB_CARGO = Abs(IMPORTE_DIF)
                  RB_DESCRIPCION_DIF = DES_CTA_CONTRA
                  GoSub ADI_DIF
              ElseIf IMPORTE_DIF > 0 Then
                  RB_CTACONT_DIF = DES_COD_FAVOR
                  RB_ABANO = Abs(IMPORTE_DIF)
                  RB_DESCRIPCION_DIF = DES_CTA_FAVOR
                  GoSub ADI_DIF
              End If
           ElseIf llave_rep01!ALL_CODTRA = 2720 Then
              PU_TIPMOV = 10
              pu_codcia = llave_rep01!all_CODCIA
              PU_NUMSER = llave_rep01!all_numser_c
              PU_FBG = llave_rep01!ALL_FBG
              PU_NUMFAC = llave_rep01!all_numfac_c
              SQ_OPER = 1
              LEER_FAR_LLAVE
              far_llave.MoveLast
              If Not far_llave.EOF Then
                If Format(far_llave!FAR_fecha_compra, "dd/mm") = "31/12" Then
                  W_IMPORTE = redondea(val(llave_rep01!ALL_IMPORTE) * JALAR_TC(far_llave!FAR_fecha_compra, 1))
                Else
                  W_IMPORTE = redondea(val(llave_rep01!ALL_IMPORTE) * JALAR_TC(far_llave!FAR_fecha_compra, 3))
                End If
                WP_DOC_PROV = redondea(val(llave_rep01!ALL_IMPORTE) * JALAR_TC(llave_rep01!ALL_FECHA_CAN, llave_rep01!ALL_SIGNO_CCM))
                'VERIFICA SI EXISTE DIF. T.C.
                IMP_CONTAB = WP_DOC_PROV 'redondea((WP_DOC_PROV - WP_NOTA_PROV) * JALAR_TC(llave_rep01!all_fecha_can, llave_rep01!all_signo_ccm))
                IMPORTE_DIF = W_IMPORTE - IMP_CONTAB
                ' ES PERDIDA
                If IMPORTE_DIF > 0 Then
                   RB_CTACONT_DIF = DES_COD_CONTRA
                   RB_CARGO = Abs(IMPORTE_DIF)
                   RB_DESCRIPCION_DIF = DES_CTA_CONTRA
                   GoSub ADI_DIF
                ElseIf IMPORTE_DIF < 0 Then
                   RB_CTACONT_DIF = DES_COD_FAVOR
                   RB_ABANO = Abs(IMPORTE_DIF)
                   RB_DESCRIPCION_DIF = DES_CTA_FAVOR
                   GoSub ADI_DIF
                End If
              End If
           ElseIf llave_rep01!ALL_CODTRA = 2735 Then
              WP_NOTAC = 0
              WP_PROV = 0
              PS_REP03(0) = llave_rep01!all_CODCIA
              PS_REP03(1) = "P"
              PS_REP03(2) = llave_rep01!ALL_CODCLIE
              PS_REP03(3) = 0
              PS_REP03(4) = Nulo_Valor0(llave_rep01!ALL_NUMGUIA) ' RELACION CON EL MISMO "NUMDOC" ORIGINAL DEL DOCUMCNET
              llave_rep03.Requery
              WP_NOTA_PROV = 0
              WP_DOC_PROV = 0
              If Not llave_rep03.EOF Then
                 Do Until llave_rep03.EOF
                   WP_NOTAC = WP_NOTAC + redondea(Abs(val(llave_rep03!CAA_IMPORTE)) * JALAR_TC(llave_rep03!CAA_FECHA_COBRO, 3))
                   WP_NOTA_PROV = WP_NOTA_PROV + Abs(val(llave_rep03!CAA_IMPORTE))
                   llave_rep03.MoveNext
                 Loop
                 PS_REP02(0) = llave_rep01!all_CODCIA
                 PS_REP02(1) = llave_rep01!all_numser_c
                 PS_REP02(2) = llave_rep01!all_numfac_c
                 PS_REP02(3) = llave_rep01!ALL_CODCLIE
                 PS_REP02(4) = "P"
                 llave_rep02.Requery
                 'Print llave_rep01!ALL_IMPORTE
                 If Not llave_rep02.EOF Then
                   WP_PROV = redondea(llave_rep02!ALL_IMPORTE_AMORT * JALAR_TC(llave_rep02!ALL_FECHA_SUNAT, 3))
                   WP_DOC_PROV = llave_rep02!ALL_IMPORTE_AMORT
                 End If
                 W_IMPORTE = WP_PROV - WP_NOTAC
                 ' VERIFICA SI EXISTE DIF. T.C.
                 IMP_CONTAB = redondea((WP_DOC_PROV - WP_NOTA_PROV) * JALAR_TC(llave_rep01!ALL_FECHA_CAN, llave_rep01!ALL_SIGNO_CCM))
                 IMPORTE_DIF = W_IMPORTE - IMP_CONTAB
                 If IMPORTE_DIF < 0 Then
                    RB_CTACONT_DIF = DES_COD_CONTRA
                    RB_CARGO = Abs(IMPORTE_DIF)
                    RB_DESCRIPCION_DIF = DES_CTA_CONTRA
                    GoSub ADI_DIF
                 ElseIf IMPORTE_DIF > 0 Then
                    RB_CTACONT_DIF = DES_COD_FAVOR
                    RB_ABANO = Abs(IMPORTE_DIF)
                    RB_DESCRIPCION_DIF = DES_CTA_FAVOR
                    GoSub ADI_DIF
                 End If
              Else
                 PS_REP02(0) = llave_rep01!all_CODCIA
                 PS_REP02(1) = llave_rep01!all_numser_c
                 PS_REP02(2) = llave_rep01!all_numfac_c
                 PS_REP02(3) = llave_rep01!ALL_CODCLIE
                 PS_REP02(4) = "P"
                 llave_rep02.Requery
                 If llave_rep02.EOF Then
                   MsgBox "NO TIENE VALOR DE PROVISION :" & RB_CONCEPTO & " " & llave_rep01!all_numfac_c
                 End If
                 If Not llave_rep02.EOF Then
                    If llave_rep01!ALL_TIPDOC = "PV" Then ' solo para provicion de CTS
                      IMP_CONTAB = redondea(val(llave_rep01!ALL_IMPORTE) * llave_rep01!ALL_TIPO_CAMBIO)
                    ElseIf llave_rep01!ALL_TIPDOC = "RC" Then ' Entregas a rendir cuentas cargo de bancos
                      IMP_CONTAB = redondea(val(llave_rep01!ALL_IMPORTE) * JALAR_TC(llave_rep02!ALL_FECHA_SUNAT, -1))
                    Else
                      IMP_CONTAB = redondea(val(llave_rep01!ALL_IMPORTE) * JALAR_TC(llave_rep01!ALL_FECHA_CAN, llave_rep01!ALL_SIGNO_CCM))
                    End If
                    If llave_rep01!ALL_TIPDOC = "AL" Then ' alquilees
                      W_IMPORTE = redondea(val(llave_rep01!ALL_IMPORTE) * JALAR_TC(llave_rep02!ALL_FECHA_SUNAT, 4))
                    ElseIf llave_rep01!ALL_TIPDOC = "RC" Then ' solo Rendir cuentas
                      W_IMPORTE = redondea(val(llave_rep01!ALL_IMPORTE) * JALAR_TC(llave_rep01!ALL_FECHA_CAN, llave_rep01!ALL_SIGNO_CCM))
                    ElseIf llave_rep01!ALL_TIPDOC = "PT" Then ' solo prestamos de terceros
                      W_IMPORTE = redondea(val(llave_rep01!ALL_IMPORTE) * JALAR_TC(llave_rep02!ALL_FECHA_SUNAT, 1))
                    ElseIf llave_rep01!ALL_TIPDOC = "PV" Then ' solo para provicion de CTS
                      If llave_rep02!ALL_MONEDA_CLI = "S" Then
                         W_IMPORTE = redondea(val(llave_rep02!ALL_IMPORTE_AMORT))
                      Else
                         W_IMPORTE = redondea(val(llave_rep01!ALL_IMPORTE_AMORT) * JALAR_TC(llave_rep02!ALL_FECHA_SUNAT, 3))
                      End If
                    Else
                      W_IMPORTE = redondea(val(llave_rep01!ALL_IMPORTE) * JALAR_TC(llave_rep02!ALL_FECHA_SUNAT, 3))
                    End If
                  ' Print llave_rep01!ALL_IMPORTE_AMORT
                  If llave_rep01!ALL_TIPDOC = "PV" Then
                    IMPORTE_DIF = IMP_CONTAB - W_IMPORTE
                  Else
                    IMPORTE_DIF = W_IMPORTE - IMP_CONTAB
                  End If
                  If IMPORTE_DIF < 0 Then
                    RB_CTACONT_DIF = DES_COD_CONTRA
                    RB_CARGO = Abs(IMPORTE_DIF)
                    RB_DESCRIPCION_DIF = DES_CTA_CONTRA
                    GoSub ADI_DIF
                  ElseIf IMPORTE_DIF > 0 Then
                    RB_CTACONT_DIF = DES_COD_FAVOR
                    RB_ABANO = Abs(IMPORTE_DIF)
                    RB_DESCRIPCION_DIF = DES_CTA_FAVOR
                    GoSub ADI_DIF
                  End If
                 End If
              End If
           Else
              W_IMPORTE = redondea(val(llave_rep01!ALL_IMPORTE) * JALAR_TC(llave_rep01!ALL_FECHA_CAN, llave_rep01!ALL_SIGNO_CCM))
           End If
 
    Else
       W_IMPORTE = val(llave_rep01!ALL_IMPORTE)
       ' VERIFICA SI EXISTE DIF. T.C.
       If llave_rep01!ALL_CODTRA = 5318 And ccm_llave!CCM_MONEDA = "D" Then
         WP_DOC_PROV = val(llave_rep01!ALL_IMPORTE_AMORT)
         IMP_CONTAB = redondea((WP_DOC_PROV) * JALAR_TC(llave_rep01!ALL_FECHA_CAN, llave_rep01!ALL_SIGNO_CCM))
         IMPORTE_DIF = W_IMPORTE - IMP_CONTAB
         ' CAMBIO DE SIGNO <>
         If llave_rep01!ALL_SIGNO_CCM = -1 Then
            If IMPORTE_DIF < 0 Then
                  RB_CTACONT_DIF = DES_COD_CONTRA
                  RB_CARGO = Abs(IMPORTE_DIF)
                  RB_DESCRIPCION_DIF = DES_CTA_CONTRA
                  GoSub ADI_DIF
            ElseIf IMPORTE_DIF > 0 Then
                  RB_CTACONT_DIF = DES_COD_FAVOR
                  RB_ABANO = Abs(IMPORTE_DIF)
                  RB_DESCRIPCION_DIF = DES_CTA_FAVOR
                  GoSub ADI_DIF
            End If
         Else
            If IMPORTE_DIF > 0 Then
                  RB_CTACONT_DIF = DES_COD_CONTRA
                  RB_CARGO = Abs(IMPORTE_DIF)
                  RB_DESCRIPCION_DIF = DES_CTA_CONTRA
                  GoSub ADI_DIF
            ElseIf IMPORTE_DIF < 0 Then
                  RB_CTACONT_DIF = DES_COD_FAVOR
                  RB_ABANO = Abs(IMPORTE_DIF)
                  RB_DESCRIPCION_DIF = DES_CTA_FAVOR
                  GoSub ADI_DIF
            End If
         End If
      End If
     End If
     If llave_rep01!ALL_CODTRA = 2735 And llave_rep01!ALL_moneda_ccm <> "D" Then
     '     If llave_rep01!ALL_IMPORTE = 161 Then Stop
           PS_REP02(0) = llave_rep01!all_CODCIA
           PS_REP02(1) = llave_rep01!all_numser_c
           PS_REP02(2) = llave_rep01!all_numfac_c
           PS_REP02(3) = llave_rep01!ALL_CODCLIE
           PS_REP02(4) = "P"
           llave_rep02.Requery
           If llave_rep02.EOF Then
              MsgBox "NO TIENE VALOR DE PROVISION :" & RB_CONCEPTO & " " & llave_rep01!all_numfac_c
           End If
           If Not llave_rep02.EOF Then
             If llave_rep02!ALL_MONEDA_CLI = "D" Then
               'Print llave_rep01!ALL_IMPORTE_AMORT
                IMP_CONTAB = redondea(redondea(val(llave_rep01!ALL_IMPORTE) / val(llave_rep01!ALL_TIPO_CAMBIO)) * JALAR_TC(llave_rep02!ALL_FECHA_SUNAT, 3))
                W_IMPORTE = redondea(redondea(val(llave_rep01!ALL_IMPORTE) / val(llave_rep01!ALL_TIPO_CAMBIO)) * val(llave_rep01!ALL_TIPO_CAMBIO)) ' JALAR_TC(llave_rep01!all_FECHA_CAN, llave_rep01!ALL_SIGNO_CCM))
                'W_IMPORTE = redondea(redondea(Val(llave_rep01!ALL_IMPORTE) / Val(llave_rep01!ALL_TIPO_CAMBIO)) * JALAR_TC(llave_rep01!all_FECHA_CAN, llave_rep01!ALL_SIGNO_CCM))
                IMPORTE_DIF = IMP_CONTAB - W_IMPORTE
                W_IMPORTE = IMP_CONTAB
                If IMPORTE_DIF < 0 Then
                  RB_CTACONT_DIF = DES_COD_CONTRA
                  RB_CARGO = Abs(IMPORTE_DIF)
                  RB_DESCRIPCION_DIF = DES_CTA_CONTRA
                  GoSub ADI_DIF
                ElseIf IMPORTE_DIF > 0 Then
                 RB_CTACONT_DIF = DES_COD_FAVOR
                 RB_ABANO = Abs(IMPORTE_DIF)
                 RB_DESCRIPCION_DIF = DES_CTA_FAVOR
                 GoSub ADI_DIF
                End If
             End If
           End If
     End If
     'End If
'     If W_IMPORTE = 12.49 Then Stop
  'Print llave_rep01!all_FECHA_CAN
 If llave_rep01!ALL_SIGNO_CCM = 1 Then ' VA PARA COLUMNA DE CARGO
     If W_IMPORTE < 0 Then
        RB_ABANO = Abs(W_IMPORTE)
     Else
        RB_CARGO = W_IMPORTE
     End If
  Else  ' VA PARA COLUMNA DE ABONO
     If W_IMPORTE < 0 Then
       RB_CARGO = Abs(W_IMPORTE)
     Else
       RB_ABANO = W_IMPORTE
     End If
  End If
  xlR.Cells(f1, 1) = RB_CTACONT
  xlR.Cells(f1, 2) = RB_DESCRIPCION
  xlR.Cells(f1, 3) = RB_COMPRO
  xlR.Cells(f1, 4) = RB_SECUENCIA
  xlR.Cells(f1, 5) = RB_TIPO
  xlR.Cells(f1, 6) = RB_NUMSER_C
  xlR.Cells(f1, 7) = RB_NUMFAC_C
  xlR.Cells(f1, 8) = RB_CONCEPTO
  'If RB_ABANO <> 0 And RB_CARGO <> 0 Then Stop
  xlR.Cells(f1, 9) = RB_ABANO
  xlR.Cells(f1, 10) = RB_CARGO
  xlR.Cells(f1, 11) = RB_NOMCORTO
  xlR.Cells(f1, 12) = val(llave_rep01!ALL_IMPORTE)
  xlR.Cells(f1, 13) = RB_CTA
  
pasa:
 llave_rep01.MoveNext
Loop
If wpasa_rep = 1 Then
 GoTo OTRO_PASE
End If

wranF = "A" & 1 & ":P" & f1
xlR.APPLICATION.Worksheets("Detalle").Range(wranF).Sort Key1:=xlR.APPLICATION.Worksheets("Detalle").Range("A1")

fin_filas = f1


CTACONT = -1 'llave_rep01!CLI_CUENTA_CONTAB
f1 = 0
fila = 5
RCRYSTAL.ProgBar.Value = 0
RCRYSTAL.ProgBar.Min = 0
If fin_filas - 1 = 0 Then
Else
RCRYSTAL.ProgBar.Max = fin_filas - 1
End If
imp_cargo = 0
imp_abono = 0
tot_cargo = 0
tot_abono = 0
RCRYSTAL.lblproceso.Caption = "Ordenando Información. . . "
DoEvents
Lini = 6
For f1 = 1 To fin_filas
  RCRYSTAL.ProgBar.Value = f1 - 1
  fila = fila + 1
 If CTACONT <> Trim(xlR.Cells(f1, 1)) Then
   If fila <> 6 Then
     xl.Cells(fila, 9) = imp_abono
     xl.Cells(fila, 10) = imp_cargo
     xl.Worksheets(1).rows(fila).RowHeight = 16
     wranF = "I" & fila & ":J" & fila
     xl.Range(wranF).Font.Bold = True
     xl.Range(wranF).Font.Size = 10
     xl.Range(wranF).Font.Size = 11
     fila = fila + 1
   End If
   xl.Cells(fila, 1) = Trim(xlR.Cells(f1, 1))
   xl.Cells(fila, 2) = Trim(xlR.Cells(f1, 2))
   wranF = "B" & fila & ":B" & fila
   xl.Range(wranF).Font.Bold = True
   CTACONT = Trim(xlR.Cells(f1, 1))
   tot_cargo = tot_cargo + imp_cargo
   tot_abono = tot_abono + imp_abono
   imp_cargo = 0
   imp_abono = 0
   fila = fila + 1
 End If
  
  xl.Cells(fila, 2) = xlR.Cells(f1, 11)
  xl.Cells(fila, 3) = "'" & Format(Trim(xlR.Cells(f1, 3)), "00000000")
  xl.Cells(fila, 4) = "'" & Format(Trim(xlR.Cells(f1, 4)), "00000")
  xl.Cells(fila, 5) = "'" & Format(xlR.Cells(f1, 5), "00")
  xl.Cells(fila, 6) = "'" & Format(Trim(xlR.Cells(f1, 6)), "000")
  xl.Cells(fila, 7) = "'" & Format(Trim(xlR.Cells(f1, 7)), "000000")
  xl.Cells(fila, 8) = Trim(xlR.Cells(f1, 8))
  xl.Cells(fila, 9) = Trim(xlR.Cells(f1, 9))
  xl.Cells(fila, 10) = Trim(xlR.Cells(f1, 10))
  xl.Cells(fila, 11) = val(xlR.Cells(f1, 12))
  xl.Cells(fila, 12) = val(xlR.Cells(f1, 13))
  xl.Cells(fila, 13) = val(xlR.Cells(f1, 1))
  imp_abono = imp_abono + val(xlR.Cells(f1, 9))
  imp_cargo = imp_cargo + val(xlR.Cells(f1, 10))
  
Next f1
RCRYSTAL.lblproceso.Caption = "Mostrnado. . . "
DoEvents
fila = fila + 1
xl.Cells(fila, 9) = imp_abono
xl.Cells(fila, 10) = imp_cargo

tot_cargo = tot_cargo + imp_cargo
tot_abono = tot_abono + imp_abono
xl.Worksheets(1).rows(fila).RowHeight = 16
wranF = "I" & fila & ":J" & fila
xl.Range(wranF).Font.Bold = True
xl.Range(wranF).Font.Size = 10
xl.Range(wranF).Font.Size = 11
Lfin = fila

DoEvents
RCRYSTAL.ProgBar.Min = 0
RCRYSTAL.ProgBar.Value = 0
RCRYSTAL.ProgBar.Max = 10

pub_mensaje = "Desea mostrar el Resumen de asiento Contable... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
 GoTo dale_informe
End If



'xl.Application.Visible = True
' RESUMEN DE ASIENTO CONTABLE
'=============================

wranF = "A" & Lini & ":S" & Lfin
xl.Sheets(1).Activate
 
'xl.Application.Visible = True

xl.APPLICATION.Worksheets(1).Range(wranF).Sort Key1:=xl.APPLICATION.Worksheets(1).Range("A1")
xl.APPLICATION.Worksheets(1).Range(wranF).Sort Key1:=xl.APPLICATION.Worksheets(1).Range("M1"), Key2:=xl.APPLICATION.Worksheets(1).Range("I1")
f1 = 4
fila = Lini
ts_codcta = Trim(Format(xl.Cells(fila, 13), "##########"))
If val(Format(xl.Cells(fila, 9), "0.000")) <> 0 Then
      xl.Cells(fila, 14) = "D"
Else
      xl.Cells(fila, 14) = "H"
End If
ts_DH = Trim(xl.Cells(fila, 14))
ts_sumaD = 0
ts_sumaH = 0
RCRYSTAL.ProgBar.Value = RCRYSTAL.ProgBar.Value + 1
For fila = Lini To Lfin
  If Trim(xl.Cells(fila, 5)) = "" Then GoTo cont_p
  
  If val(Format(xl.Cells(fila, 9), "0.000")) <> 0 Then
      xl.Cells(fila, 14) = "D"
  Else
      xl.Cells(fila, 14) = "H"
  End If
  
  If Trim(ts_codcta) <> Trim(Format(xl.Cells(fila, 13), "##########")) Or Trim(ts_DH) <> Trim(xl.Cells(fila, 14)) Then
    xl.Sheets(2).Activate
    f1 = f1 + 1
    xl.Cells(f1, 1) = ts_codcta
    xl.Cells(f1, 2) = JALA_CTA(ts_codcta)
    xl.Cells(f1, 3) = ts_sumaD ' debe
    xl.Cells(f1, 4) = ts_sumaH  ' haber
    If ts_sumaD <> 0 Then
      xl.Cells(f1, 5) = "D"
    Else
      xl.Cells(f1, 5) = "H"
    End If
    xl.Sheets(1).Activate
    ts_codcta = Trim(Format(xl.Cells(fila, 13), "##########"))
    If val(Format(xl.Cells(fila, 9), "0.000")) <> 0 Then
          xl.Cells(fila, 14) = "D"
    Else
          xl.Cells(fila, 14) = "H"
    End If
    ts_DH = Trim(xl.Cells(fila, 14))
    ts_sumaD = 0
    ts_sumaH = 0
  End If
'  xl.Application.Visible = True
  ts_sumaD = ts_sumaD + val(Format(xl.Cells(fila, 9), "0.000"))
  ts_sumaH = ts_sumaH + val(Format(xl.Cells(fila, 10), "0.000"))
  
cont_p:
Next fila
xl.Sheets(2).Activate

xl.Cells(1, 1) = WEMPRESA '
xl.Cells(2, 1) = "PERIODO: '" & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
f1 = f1 + 1
xl.Cells(f1, 1) = ts_codcta
xl.Cells(f1, 2) = JALA_CTA(ts_codcta)
xl.Cells(f1, 3) = ts_sumaD ' debe
xl.Cells(f1, 4) = ts_sumaH  ' haber
If ts_sumaD <> 0 Then
  xl.Cells(f1, 5) = "D"
Else
  xl.Cells(f1, 5) = "H"
End If
    
ts_sumaD = 0
ts_sumaH = 0
RCRYSTAL.ProgBar.Value = RCRYSTAL.ProgBar.Value + 1
'----------FIN DE LA PRIMERA PARTE-------
'***************************************
'---------------------------------------

wranF = "A" & Lini & ":S" & Lfin
xl.Sheets(1).Activate
 
'xl.Application.Visible = True
xl.APPLICATION.Worksheets(1).Range(wranF).Sort Key1:=xl.APPLICATION.Worksheets(1).Range("A1")
xl.APPLICATION.Worksheets(1).Range(wranF).Sort Key1:=xl.APPLICATION.Worksheets(1).Range("L1"), Key2:=xl.APPLICATION.Worksheets(1).Range("I1")

fila = Lini
ts_codcta = Trim(Format(xl.Cells(fila, 12), "##########"))
ts_DH = Trim(xl.Cells(fila, 14))
ts_sumaD = 0
ts_sumaH = 0
RCRYSTAL.ProgBar.Value = RCRYSTAL.ProgBar.Value + 1
For fila = Lini To Lfin
  If Trim(xl.Cells(fila, 5)) = "" Then GoTo cont_p2
  If Trim(ts_codcta) <> Trim(Format(xl.Cells(fila, 12), "##########")) Or Trim(ts_DH) <> Trim(xl.Cells(fila, 14)) Then
    xl.Sheets(2).Activate
    f1 = f1 + 1
    xl.Cells(f1, 1) = ts_codcta
    xl.Cells(f1, 2) = JALA_CTA(ts_codcta)
    xl.Cells(f1, 4) = ts_sumaD ' debe
    xl.Cells(f1, 3) = ts_sumaH  ' haber
    If ts_sumaD <> 0 Then
      xl.Cells(f1, 5) = "H" ' INVIERTEN LOS DEBES POR LOS HABER
    Else
      xl.Cells(f1, 5) = "D"
    End If
    xl.Sheets(1).Activate
    ts_codcta = Trim(Format(xl.Cells(fila, 12), "##########"))
    ts_DH = Trim(xl.Cells(fila, 14))
    ts_sumaD = 0
    ts_sumaH = 0
  End If
'  xl.Application.Visible = True
  ts_sumaD = ts_sumaD + val(Format(xl.Cells(fila, 9), "0.000"))
  ts_sumaH = ts_sumaH + val(Format(xl.Cells(fila, 10), "0.000"))
  
cont_p2:
Next fila
RCRYSTAL.ProgBar.Value = RCRYSTAL.ProgBar.Value + 1
xl.Sheets(2).Activate
xl.Cells(1, 1) = WEMPRESA '
xl.Cells(2, 1) = "PERIODO: '" & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
f1 = f1 + 1
xl.Cells(f1, 1) = ts_codcta
xl.Cells(f1, 2) = JALA_CTA(ts_codcta)
xl.Cells(f1, 4) = ts_sumaD ' debe
xl.Cells(f1, 3) = ts_sumaH  ' haber
If ts_sumaD <> 0 Then
  xl.Cells(f1, 5) = "H" ' INVIERTEN LOS DEBES POR LOS HABER
Else
  xl.Cells(f1, 5) = "D"
End If
    
ts_sumaD = 0
ts_sumaH = 0
RCRYSTAL.ProgBar.Value = RCRYSTAL.ProgBar.Value + 1

'xl.Application.Visible = True
'Stop
'----------SEGUNDA PARTE ----------------

'----------FIN DE LA SEGUNDA PARTE-------

' ORDER PARA ASIENTO DE LA 101001 CAJA
xl.Sheets(2).Activate
wranF = "A" & 5 & ":E" & f1
xl.Sheets(2).Activate
'xl.Application.Visible = True
xl.APPLICATION.Worksheets(2).Range(wranF).Sort Key1:=xl.APPLICATION.Worksheets(2).Range("A1")
xl.APPLICATION.Worksheets(2).Range(wranF).Sort Key1:=xl.APPLICATION.Worksheets(2).Range("E1"), Key2:=xl.APPLICATION.Worksheets(2).Range("A1")
' SEPARACION DE ASIENTOS DE INGRESOS Y EGRESOS
Lini = 0
fila = 5
ts_DH = Trim(xl.Cells(fila, 5))
ts_suma = 0
RCRYSTAL.ProgBar.Value = RCRYSTAL.ProgBar.Value + 1
For fila = 5 To f1 + 2
  If Trim(xl.Cells(fila, 1)) = "" Then GoTo cont_p3
  If Trim(ts_DH) <> Trim(xl.Cells(fila, 5)) Then
    wranF = "A" & fila
    xl.Range(wranF).Select
    xl.Selection.EntireRow.Insert
    xl.Selection.EntireRow.Insert
    ts_codcta = "101001"
    xl.Cells(fila, 1) = ts_codcta
    xl.Cells(fila, 2) = JALA_CTA(ts_codcta)
    xl.Cells(fila, 4) = ts_suma
    xl.Cells(fila, 3) = 0
    If ts_DH = "D" Then
      ts_DH = "H"
    Else
      ts_DH = "D"
    End If
    xl.Cells(fila, 5) = ts_DH
    fila = fila + 1
    xl.Cells(fila, 1) = "Total Egresos"
    wran1 = "C" & 5
    wran2 = "C" & fila - 1
    wranF = "C" & fila
    xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
    wran1 = "D" & 5
    wran2 = "D" & fila - 1
    wranF = "D" & fila
    xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
    
    fila = fila + 1
    
    ts_DH = Trim(xl.Cells(fila, 5))
    Lini = fila
    ts_suma = 0
  End If
  ts_suma = ts_suma + val(Format(xl.Cells(fila, 3), "0.000")) + val(Format(xl.Cells(fila, 4), "0.000"))
cont_p3:
Next fila
RCRYSTAL.ProgBar.Value = RCRYSTAL.ProgBar.Value + 1
'fila = fila + 1
ts_codcta = "101001"
xl.Cells(fila, 1) = ts_codcta
xl.Cells(fila, 2) = JALA_CTA(ts_codcta)
xl.Cells(fila, 3) = ts_suma
xl.Cells(fila, 4) = 0
If ts_DH = "D" Then
ts_DH = "H"
Else
ts_DH = "D"
End If
xl.Cells(fila, 5) = ts_DH
fila = fila + 1
xl.Cells(fila, 1) = "Total Ingresos"
wran1 = "C" & Lini
wran2 = "C" & fila - 1
wranF = "C" & fila
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
wran1 = "D" & Lini
wran2 = "D" & fila - 1
wranF = "D" & fila
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
fila = fila + 1
ts_DH = Trim(xl.Cells(fila, 5))

ts_suma = 0
RCRYSTAL.ProgBar.Value = RCRYSTAL.ProgBar.Value + 1
RCRYSTAL.ProgBar.Value = RCRYSTAL.ProgBar.Value + 1
RCRYSTAL.ProgBar.Value = RCRYSTAL.ProgBar.Value + 1
dale_informe:
xl.Sheets(1).Activate

fila = fila + 1
xl.Cells(fila, 8) = "TOTAL GENERAL = "
xl.Cells(fila, 9) = tot_abono
xl.Cells(fila, 10) = tot_cargo
xl.Worksheets(1).rows(fila).RowHeight = 18
wranF = "I" & fila & ":J" & fila
xl.Range(wranF).Font.Bold = True
xl.Range(wranF).Font.Size = 10
xl.Range(wranF).Font.Size = 11


If WEMPRESA = "" Then
  WEMPRESA = Trim(par_llave!PAR_NOMBRE)
Else
  WEMPRESA = Trim(GEN!GEN_NOMBRE) & " " & WEMPRESA
End If

xl.Cells(1, 1) = WEMPRESA '
xl.Cells(2, 1) = Trim(retra_llave!TRA_DESCRIPCION)
xl.Cells(3, 1) = "'" & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
xl.Cells(2, 4) = UCase(Trim(RCRYSTAL.lblbanco))

If chepasa.Value = 1 Then
    xl.Sheets(2).Activate
 '   xl.Application.Visible = True
    ASIENTO_MOVICONT xl, 3
End If

  
xl.APPLICATION.Visible = True
RCRYSTAL.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
'xlR.Application.Worksheets(1).Range(wranF).Sort Key1:=xl.Application.Worksheets(1).Range("E1"), Key2:=xl.Application.Worksheets("Hoja1").Range("F1"), Key3:=xl.Application.Worksheets("Hoja1").Range("G1")
xlR.DisplayAlerts = False
xlR.Worksheets(1).Protect PUB_CLAVE
xlR.Workbooks(1).Close
xl.DisplayAlerts = False
xl.Worksheets(1).Protect ""
DoEvents
RCRYSTAL.lblproceso.Visible = False
RCRYSTAL.ProgBar.Visible = False
Set xl = Nothing
Set xlR = Nothing
Screen.MousePointer = 0
RCRYSTAL.Pantalla.Enabled = True
RCRYSTAL.Pantalla.Caption = "Por &Pantalla"
RCRYSTAL.lblproceso.Visible = False
Pantalla.Enabled = True
cmdCerrar.Enabled = True

Exit Sub



CANCELA:
  RCRYSTAL.Pantalla.Enabled = True
  RCRYSTAL.Pantalla.Caption = "Por &Pantalla"
  RCRYSTAL.lblproceso.Visible = False
  RCRYSTAL.ProgBar.Visible = False
  Pantalla.Enabled = True
  cmdCerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
  Set xl = Nothing
  If xlR Is Nothing Then
  Else
   xlR.APPLICATION.Visible = True
  End If
  Set xlR = Nothing
  Screen.MousePointer = 0
Exit Sub

WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  If xlR Is Nothing Then
     Set xlR = CreateObject("Excel.Application")
  End If
  
  
  DoEvents
  xl.Workbooks.Open PUB_RUTA_OTRO & "RBANCOS.xls", 0, True, 4, "", ""
  xlR.Workbooks.Open PUB_RUTA_OTRO & "RBANCOSTEM.xls", 0, True, 4, "", ""
Return

TOTAL_DIA:
Return

ADI_DIF:
'If (RB_ABANO + RB_CARGO) = 360 Then Stop

  xlR.Cells(f1, 1) = RB_CTACONT_DIF
  xlR.Cells(f1, 2) = RB_DESCRIPCION_DIF
  xlR.Cells(f1, 3) = RB_COMPRO
  xlR.Cells(f1, 4) = RB_SECUENCIA
  xlR.Cells(f1, 5) = RB_TIPO
  xlR.Cells(f1, 6) = RB_NUMSER_C
  xlR.Cells(f1, 7) = RB_NUMFAC_C
  xlR.Cells(f1, 8) = RB_CONCEPTO
  xlR.Cells(f1, 9) = RB_CARGO
  xlR.Cells(f1, 10) = RB_ABANO
  xlR.Cells(f1, 11) = RB_NOMCORTO
  xlR.Cells(f1, 12) = val(llave_rep01!ALL_IMPORTE)
  xlR.Cells(f1, 13) = RB_CTA
  RB_CARGO = 0
  RB_ABANO = 0
  f1 = f1 + 1
Return

FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
 GoTo CANCELA
 Resume Next

End Sub

Public Sub MOVI_BANCO()
'On Error GoTo FINTODO
Dim RB_CTACONT As String
Dim WP_NOTAC As Currency
Dim WP_PROV As Currency
Dim W_IMPORTE As Currency
Dim wcheque As Currency
Dim wsalmacenes As String
Dim WTC As Currency
Dim WCIA1 As String * 2
Dim WCIA2 As String * 2
Dim WCIA3 As String * 2
Dim WCIA4 As String * 2
Dim WSCODART As Currency
Dim FF1 As Integer
Dim wsum_abono  As Currency
Dim wsum_cargo  As Currency
Dim wfecha
Dim wdocumento  As String
Dim wdocserie  As String
Dim wglosa As String
Dim wciu  As String
Dim wnumfac As String
Dim s_total_abono As Currency
Dim s_total_cargo As Currency
If val(txt_key.Text) <= 0 Then
  MsgBox "Ingrese banco para procesar...", 48, Pub_Titulo
  Azul txt_key, txt_key
  Exit Sub
End If
SQ_OPER = 1
PUB_CODBAN = val(txt_key.Text)
LEER_CCM_LLAVE
If ccm_llave.EOF Then
  MsgBox "Banco no procede...", 48, Pub_Titulo
  Azul txt_key, txt_key
  Exit Sub
End If

Pantalla.Enabled = False
cmdCerrar.Enabled = False
DoEvents
'FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
WCIA1 = ""
WCIA2 = ""
WCIA3 = ""
WCIA4 = ""

wsalmacenes = ""
If LK_EMP <> "3AA" Then
 WCIA1 = LK_CODCIA
 GoTo OTRO
End If
For fila = 0 To LISCIA.ListCount - 1
 LISCIA.ListIndex = fila
 If LISCIA.Selected(fila) Then
    PSPAR_MULTI(0) = Left(LISCIA.Text, 2)
    par_multi.Requery
    wsalmacenes = wsalmacenes + Trim(par_multi!par_nombre_corto) & " - "
 End If
Next fila
If wsalmacenes <> "" Then
  wsalmacenes = Mid(wsalmacenes, 1, Len(wsalmacenes) - 3)
Else
  wsalmacenes = par_llave!PAR_NOMBRE
End If

For fila = 0 To LISCIA.ListCount - 1
LISCIA.ListIndex = fila
If LISCIA.Selected(fila) Then
    If Trim(WCIA1) = "" Then
     WCIA1 = Left(LISCIA.Text, 2)
    ElseIf Trim(WCIA2) = "" Then
     WCIA2 = Left(LISCIA.Text, 2)
    ElseIf Trim(WCIA3) = "" Then
     WCIA3 = Left(LISCIA.Text, 2)
    ElseIf Trim(WCIA4) = "" Then
     WCIA4 = Left(LISCIA.Text, 2)
    End If
End If
Next fila

If Trim(WCIA1) = "" And Trim(WCIA2) = "" And Trim(WCIA3) = "" And Trim(WCIA4) = "" Then
  For fila = 0 To LISCIA.ListCount - 1
    LISCIA.ListIndex = fila
    If fila = 0 Then
       WCIA1 = Left(LISCIA.Text, 2)
    End If
    If fila = 1 Then
       WCIA2 = Left(LISCIA.Text, 2)
    End If
    If fila = 2 Then
       WCIA3 = Left(LISCIA.Text, 2)
    End If
    If fila = 3 Then
       WCIA4 = Left(LISCIA.Text, 2)
    End If
  Next fila
End If
OTRO:
If SON_FECHAS = False Then Exit Sub
GoSub WEXCEL
DoEvents
pub_cadena = "SELECT * FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_NUMSER_C = ? AND ALL_NUMFAC_C = ?  AND ALL_CODCLIE = ? AND ALL_CP = ? AND ALL_FLAG_EXT <> 'E'"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = 0
PS_REP01(3) = 0
PS_REP01(4) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT CAA_SALDO_CAR, CAA_IMPORTE, CAA_FECHA_COBRO FROM CARACU WHERE CAA_CODCIA = ? AND CAA_CP = ? AND CAA_CODCLIE = ? AND CAA_SERDOC = ? AND CAA_NUMDOC = ?  AND CAA_NOTA = 'C'  AND CAA_ESTADO <> 'E' "
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
PS_REP03(0) = 0
PS_REP03(1) = 0
PS_REP03(2) = 0
PS_REP03(3) = 0
PS_REP03(4) = 0
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ? OR ALL_CODCIA = ? OR ALL_CODCIA = ? )  AND (ALL_CODBAN = ?) AND (ALL_FECHA_CAN >= ? AND ALL_FECHA_CAN <= ? ) AND all_codban <> 0 and all_codtra <> 2401  and all_codtra <> 2725 AND all_signo_ccm  <> 0 AND ALL_FLAG_EXT <> 'E'  ORDER BY ALL_FECHA_CAN, ALL_SIGNO_CCM DESC,ALL_CHENUM"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
PS_REP02(1) = 0
PS_REP02(2) = 0
PS_REP02(3) = 0
PS_REP02(4) = 0
PS_REP02(5) = LK_FECHA_DIA
PS_REP02(6) = LK_FECHA_DIA
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PS_REP02(0) = WCIA1
PS_REP02(1) = WCIA2
PS_REP02(2) = WCIA3
PS_REP02(3) = WCIA4
PS_REP02(4) = val(txt_key.Text)
PS_REP02(5) = txtCampo1.Text
PS_REP02(6) = txtCampo2.Text



DoEvents
'xl.Cells(1, 1) = Trim(GEN!GEN_NOMBRE)
xl.Cells(2, 1) = "REGISTRO DE BANCO " + UCase(wsalmacenes)
xl.Cells(4, 1) = "Fecha del : " & txtCampo1.Text & "  Al   " & txtCampo2.Text
xl.Cells(5, 1) = "BANCO : "
xl.Cells(5, 3) = Trim(ccm_llave!CCM_NOMBRE)
If Not ccm_llave.EOF Then
    If ccm_llave!CCM_MONEDA = "D" Then
     xl.Cells(7, 9) = "US$."
     xl.Cells(7, 10) = "US$."
    Else
     xl.Cells(7, 9) = "S/."
     xl.Cells(7, 10) = "S/."
    End If
    If cheflag.Value = 1 Then
      xl.Cells(7, 9) = "M.N."
      xl.Cells(7, 10) = "M.N."
    End If
End If
f1 = 6  'Fila Inicial
'PS_REP02(0) = WCIA1  ''LK_CODCIA
llave_rep02.Requery
If llave_rep02.EOF Then
  MsgBox "No existe Movimiento", 48, Pub_Titulo
  GoTo CANCELA
End If
RCRYSTAL.ProgBar.Min = 0
RCRYSTAL.ProgBar.Value = 0
RCRYSTAL.ProgBar.Max = llave_rep02.RowCount
RCRYSTAL.ProgBar.Visible = True
DoEvents
RCRYSTAL.lblproceso.Visible = True
RCRYSTAL.lblproceso.Caption = "Procesando Información. . . "
DoEvents
wcheque = val(llave_rep02!all_chenum)
f1 = 7
s_total_abono = 0
s_total_cargo = 0
Do Until llave_rep02.EOF
    RCRYSTAL.ProgBar.Value = RCRYSTAL.ProgBar.Value + 1
    If wcheque <> val(llave_rep02!all_chenum) Then
        f1 = f1 + 1
        xl.Cells(f1, 1) = wfecha
        xl.Cells(f1, 2) = ""
        xl.Cells(f1, 3) = "'" & Format(wcheque, "00000000")
        xl.Cells(f1, 4) = wnumfac
        xl.Cells(f1, 5) = wciu
        xl.Cells(f1, 6) = val(wdocserie)
        xl.Cells(f1, 7) = val(wdocumento)
        xl.Cells(f1, 8) = wglosa
        xl.Cells(f1, 9) = val(wsum_abono)
        xl.Cells(f1, 10) = val(wsum_cargo)
        s_total_abono = s_total_abono + val(wsum_abono)
        s_total_cargo = s_total_cargo + val(wsum_cargo)
        wcheque = val(llave_rep02!all_chenum)
        wsum_cargo = 0
        wsum_abono = 0
     End If
     W_IMPORTE = 0
     'If Val(llave_rep02!all_chenum) = 1 Then Stop
'     If Val(llave_rep02!all_chenum) = 830211 Then Stop
     If llave_rep02!ALL_moneda_ccm = "D" Then
         If cheflag.Value = 1 Then
           If llave_rep02!ALL_CODTRA = 2748 Then
                W_IMPORTE = redondea(val(llave_rep02!ALL_IMPORTE) * JALAR_TC(llave_rep02!ALL_FECHA_SUNAT, 3))
           ElseIf llave_rep02!ALL_CODTRA = 5714 And llave_rep02!ALL_SIGNO_CCM = 1 And llave_rep02!ALL_IMPORTE_AMORT <> 0 Then
                W_IMPORTE = redondea(val(llave_rep02!ALL_IMPORTE_AMORT))
           ElseIf llave_rep02!ALL_CODTRA = 5318 Then
              If val(llave_rep02!ALL_IMPORTE) <> val(llave_rep02!ALL_IMPORTE_AMORT) Then
                W_IMPORTE = redondea(val(llave_rep02!ALL_IMPORTE))
              Else
                W_IMPORTE = redondea(val(llave_rep02!ALL_IMPORTE) * JALAR_TC(llave_rep02!ALL_FECHA_CAN, -1))
              End If
           ElseIf llave_rep02!ALL_CODTRA = 2720 Then
              PU_TIPMOV = 10
              pu_codcia = llave_rep02!all_CODCIA
              PU_NUMSER = llave_rep02!all_numser_c
              PU_FBG = llave_rep02!ALL_FBG
              PU_NUMFAC = llave_rep02!all_numfac_c
              SQ_OPER = 1
              LEER_FAR_LLAVE
              far_llave.MoveLast
              If Not far_llave.EOF Then
              W_IMPORTE = redondea(val(llave_rep02!ALL_IMPORTE) * JALAR_TC(far_llave!FAR_fecha_compra, 3))
              End If
           ElseIf llave_rep02!ALL_CODTRA = 2735 Then
              WP_NOTAC = 0
              WP_PROV = 0
              PS_REP03(0) = llave_rep02!all_CODCIA
              PS_REP03(1) = llave_rep02!ALL_CP
              PS_REP03(2) = llave_rep02!ALL_CODCLIE
              PS_REP03(3) = 0
              PS_REP03(4) = Nulo_Valor0(llave_rep02!ALL_NUMGUIA) ' RELACION CON EL MISMO "NUMDOC" ORIGINAL DEL DOCUMCNET
              llave_rep03.Requery
              If Not llave_rep03.EOF Then
                 Do Until llave_rep03.EOF
                   WP_NOTAC = WP_NOTAC + redondea(Abs(val(llave_rep03!CAA_IMPORTE)) * JALAR_TC(llave_rep03!CAA_FECHA_COBRO, 3))
                   llave_rep03.MoveNext
                 Loop
                 PS_REP01(0) = llave_rep02!all_CODCIA
                 PS_REP01(1) = llave_rep02!all_numser_c
                 PS_REP01(2) = llave_rep02!all_numfac_c
                 PS_REP01(3) = llave_rep02!ALL_CODCLIE
                 PS_REP01(4) = "P"
                 llave_rep01.Requery
                 If Not llave_rep01.EOF Then
                   WP_PROV = redondea(llave_rep01!ALL_IMPORTE_AMORT * JALAR_TC(llave_rep01!ALL_FECHA_SUNAT, 3))
                 End If
                 W_IMPORTE = WP_PROV - WP_NOTAC
              Else
                 PS_REP01(0) = llave_rep02!all_CODCIA
                 PS_REP01(1) = llave_rep02!all_numser_c
                 PS_REP01(2) = llave_rep02!all_numfac_c
                 PS_REP01(3) = llave_rep02!ALL_CODCLIE
                 PS_REP01(4) = "P"
                 llave_rep01.Requery
                 If Not llave_rep01.EOF Then
                    If llave_rep02!ALL_TIPDOC = "AL" Then
                       W_IMPORTE = redondea(val(llave_rep02!ALL_IMPORTE) * JALAR_TC(llave_rep01!ALL_FECHA_SUNAT, 4))
                    ElseIf llave_rep02!ALL_TIPDOC = "RC" Then ' solo Rendir cuentas
                       W_IMPORTE = redondea(val(llave_rep02!ALL_IMPORTE) * JALAR_TC(llave_rep02!ALL_FECHA_CAN, llave_rep02!ALL_SIGNO_CCM))
                    ElseIf llave_rep02!ALL_TIPDOC = "PT" Then
                       W_IMPORTE = redondea(val(llave_rep02!ALL_IMPORTE) * JALAR_TC(llave_rep01!ALL_FECHA_SUNAT, 1))
                    ElseIf llave_rep02!ALL_TIPDOC = "PV" Then
                       If llave_rep02!ALL_MONEDA_CLI = "S" Then
                         W_IMPORTE = redondea(val(llave_rep01!ALL_IMPORTE_AMORT))
                       Else
                         W_IMPORTE = redondea(val(llave_rep02!ALL_IMPORTE_AMORT) * JALAR_TC(llave_rep02!ALL_FECHA_SUNAT, 3))
                       End If
                       '  W_IMPORTE = redondea(Val(llave_rep01!ALL_IMPORTE_AMORT))
                    Else
                       W_IMPORTE = redondea(val(llave_rep02!ALL_IMPORTE) * JALAR_TC(llave_rep01!ALL_FECHA_SUNAT, 3))
                    End If
                 End If
              End If
           Else
              W_IMPORTE = redondea(val(llave_rep02!ALL_IMPORTE) * JALAR_TC(llave_rep02!ALL_FECHA_CAN, llave_rep02!ALL_SIGNO_CCM))
           End If
         Else
             W_IMPORTE = val(llave_rep02!ALL_IMPORTE)
         End If
     Else
         If cheflag.Value = 0 Then
           If llave_rep02!ALL_CODTRA = 5318 And ccm_llave!CCM_MONEDA <> llave_rep02!ALL_moneda_ccm Then
               W_IMPORTE = val(llave_rep02!ALL_IMPORTE_AMORT)
           Else
             W_IMPORTE = val(llave_rep02!ALL_IMPORTE)
           End If
         Else
            If llave_rep02!ALL_CODTRA = 2735 Then
                PS_REP01(0) = llave_rep02!all_CODCIA
                PS_REP01(1) = llave_rep02!all_numser_c
                PS_REP01(2) = llave_rep02!all_numfac_c
                PS_REP01(3) = llave_rep02!ALL_CODCLIE
                PS_REP01(4) = "P"
                llave_rep01.Requery
                If llave_rep01.EOF Then
                   MsgBox "NO TIENE VALOR DE PROVISION :" & llave_rep02!all_numfac_c
                End If
                If Not llave_rep01.EOF Then
                  If llave_rep01!ALL_MONEDA_CLI = "D" Then
                     W_IMPORTE = redondea(redondea(val(llave_rep02!ALL_IMPORTE) / val(llave_rep02!ALL_TIPO_CAMBIO)) * JALAR_TC(llave_rep01!ALL_FECHA_SUNAT, 3))
                     'W_IMPORTE = redondea(redondea(Val(llave_rep01!ALL_IMPORTE) / Val(llave_rep01!ALL_TIPO_CAMBIO)) * JALAR_TC(llave_rep01!all_FECHA_CAN, llave_rep01!all_signo_ccm))
                  Else
                     W_IMPORTE = val(llave_rep02!ALL_IMPORTE)
                  End If
                End If
            Else
               W_IMPORTE = val(llave_rep02!ALL_IMPORTE)
            End If
         End If
     End If
  '   If W_IMPORTE = 9.06 Then Stop
     If llave_rep02!ALL_SIGNO_CCM = 1 Then
       wsum_abono = wsum_abono + W_IMPORTE
     End If
     If llave_rep02!ALL_SIGNO_CCM = -1 Then
       wsum_cargo = wsum_cargo + W_IMPORTE
     End If
     wfecha = "'" & Format(llave_rep02!ALL_FECHA_CAN, "dd/mm/yy")
     wdocumento = llave_rep02!all_numfac_c
     wdocserie = llave_rep02!all_numser_c
     wglosa = Trim(llave_rep02!all_concepto)
     wciu = "'" & Format(llave_rep02!all_CODCIA, "00")
     wnumfac = llave_rep02!all_numfac
     If llave_rep02!ALL_CODTRA = 2748 Then
      SQ_OPER = 1
      pu_cp = llave_rep02!ALL_CP
      pu_codclie = llave_rep02!ALL_CODCLIE
      pu_codcia = llave_rep02!all_CODCIA
      LEER_CLI_LLAVE
      If Not cli_llave.EOF Then wglosa = Trim(cli_llave!CLI_NOMBRE)
     End If
   
llave_rep02.MoveNext
Loop
' ultimo registro
f1 = f1 + 1
xl.Cells(f1, 1) = wfecha
xl.Cells(f1, 2) = ""
xl.Cells(f1, 3) = "'" & Format(wcheque, "00000000")
xl.Cells(f1, 4) = wnumfac
xl.Cells(f1, 5) = wciu
xl.Cells(f1, 6) = wdocserie
xl.Cells(f1, 7) = wdocumento
xl.Cells(f1, 8) = wglosa
xl.Cells(f1, 9) = val(wsum_abono)
xl.Cells(f1, 10) = val(wsum_cargo)
s_total_abono = s_total_abono + val(wsum_abono)
s_total_cargo = s_total_cargo + val(wsum_cargo)
wcheque = 0
wsum_cargo = 0
wsum_abono = 0
f1 = f1 + 1
xl.Worksheets(1).rows(f1).RowHeight = 15
wranF = "H" & f1 & ":J" & f1
xl.Range(wranF).Font.Bold = True
xl.Cells(f1, 8) = "TOTAL GENERAL = "
xl.Cells(f1, 9) = s_total_abono
xl.Cells(f1, 10) = s_total_cargo

  RCRYSTAL.lblproceso.Caption = "Procesando . . .  un Momento ."
  RCRYSTAL.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  DoEvents
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect PUB_CLAVE
  xl.APPLICATION.Visible = True
  DoEvents
  RCRYSTAL.lblproceso.Visible = False
  RCRYSTAL.ProgBar.Visible = False
  Set xl = Nothing
   Screen.MousePointer = 0
  ProgBar.Visible = False
  lblproceso.Visible = False
  Pantalla.Enabled = True
  cmdCerrar.Enabled = True
Exit Sub

WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblproceso.Caption = "Abriendo , Archivo bancos.xls . . . "
  DoEvents
  xl.Workbooks.Open PUB_RUTA_OTRO & "MOVI_BANCO.xls", 0, True, 4

Return



Exit Sub
CANCELA:
  cmdCerrar.Enabled = True
  RCRYSTAL.Pantalla.Enabled = True
  RCRYSTAL.Pantalla.Caption = "Por &Pantalla"
  RCRYSTAL.lblproceso.Visible = False
  RCRYSTAL.ProgBar.Visible = False
  Pantalla.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
OJO:
If Err.Number = 70 Then
  MsgBox "Hoja de Calculo : " & wsfile1 & "  esta Abierta debe cerrar para Procesar Nuevamente ", 48, Pub_Titulo
  GoTo CANCELA
End If
Exit Sub

FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 xl.APPLICATION.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
Exit Sub

End Sub

Public Sub CHEQUEO_DESCTO()
'On Error GoTo FINTODO
Dim q_stock_val As Currency
Dim TOTAL_CLASE As Currency
Dim TOTAL_CLASE_VAL As Currency
Dim WCONCIA As Integer
Dim Wfecha_resulta As Date
Dim CHE_KARDEX As Currency
Dim WCOSPRO_SUP As Currency
Dim WTC As Currency
Dim wtotal As Currency
Dim WCIA1 As String * 2
Dim WCIA2 As String * 2
Dim WCIA3 As String * 2
Dim WCIA4 As String * 2
Dim WSCODART As Currency
Dim flag_xx As Integer
Dim ww_concepto As String
Dim ww_codcia As String * 2
Dim WS_PRECIO As Currency
Dim WW_LINEA, i
Dim ws_clave As String
Dim FF1 As Integer
Dim LETRAS(24) As String * 1
Dim WSFECHA As Date
Dim acu_val_ingresos As Currency
Dim acu_val_salidas As Currency
Dim xx_ingreso As Currency
Dim xx_salida As Currency
Dim ww_ingreso As Currency
Dim ww_salida As Currency
Dim acu_cant_dia As Currency
Dim acu_saldo As Currency
Dim acu_stock As Currency
Dim wsfile As String
Dim walterno As String
Dim wdnombre As String
Dim WD_COSPRO As Currency
Dim q_sum_calse As Currency
Dim q_sum_total As Currency
Dim q_stock As Currency
wsfile = ""
Pantalla.Enabled = False
DoEvents
'FrmImp2.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
PRO_REPORTE (1)
WCIA1 = ""
WCIA2 = ""
WCIA3 = ""
WCIA4 = ""

pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = 20  AND FAR_ESTADO <>'E' AND FAR_TOT_DESCTO <> 0 ORDER BY FAR_FECHA_COMPRA "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
DoEvents
ws_clave = PUB_CLAVE
GoSub WEXCEL
xl.Cells(1, 1) = Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia")))
xl.Cells(4, 3) = "kardex del : " & txtCampo1.Text & "  Al   " & txtCampo2.Text
f1 = 5  'Fila Inicial
PS_REP01(0) = LK_CODCIA

RCRYSTAL.lblproceso.Visible = True
RCRYSTAL.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
CHE_KARDEX = 0
DoEvents
wtotal = 0
WD_COSPRO = 0
acu_saldo = 0
WCOSPRO_SUP = 0
TOTAL_CLASE = 0
llave_rep01.Requery
Do Until llave_rep01.EOF
        TOTAL_CLASE = redondea(val(llave_rep01!FAR_SUBTOTAL) * (val(llave_rep01!FAR_TOT_DESCTO) / 100))
        If TOTAL_CLASE <> val(llave_rep01!FAR_DESCTO) Then
            Stop
            llave_rep01.Edit
            llave_rep01!FAR_DESCTO = TOTAL_CLASE
            llave_rep01.Update
            f1 = f1 + 1
            SQ_OPER = 1
            PUB_KEY = llave_rep01!far_codart
            pu_codcia = LK_CODCIA
            LEER_ART_LLAVE
            xl.Cells(f1, 1) = llave_rep01!FAR_fecha_compra
            xl.Cells(f1, 2) = art_LLAVE!ART_ALTERNO
            xl.Cells(f1, 3) = llave_rep01!far_codart
            
        End If

llave_rep01.MoveNext
Loop
  RCRYSTAL.lblproceso.Caption = "Procesando . . .  un Momento ."
  'xl.Application.Worksheets("HOJA DE COMISIONES x VENDEDOR").Range(wran1).Sort Key1:=xl.Application.Worksheets("HOJA DE COMISIONES x VENDEDOR").Range("O7")
  RCRYSTAL.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
 ' xl.Cells(2, 2) = "INFORME DE COMISIONES x VENDEDOR"
  DoEvents
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect PUB_CLAVE
  xl.APPLICATION.Visible = True
  DoEvents
  RCRYSTAL.lblproceso.Visible = False
  RCRYSTAL.ProgBar.Visible = False
  Set xl = Nothing
    Screen.MousePointer = 0
  ProgBar.Visible = False
  lblproceso.Visible = False
  Pantalla.Enabled = True
  cmdCerrar.Enabled = True
  ''Unload RCRYSTAL
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
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblproceso.Caption = "Abriendo , Archivo Comisiones.xls . . . "
  DoEvents
  xl.Workbooks.Open Left(Trim(PUB_RUTA_OTRO), 1) & ":\ADMIN\STANDAR\KARDEX_CLASES.xls", 0, True, 4

Return



Exit Sub
CANCELA:
  RCRYSTAL.Pantalla.Enabled = True
  RCRYSTAL.Pantalla.Caption = "Por &Pantalla"
  RCRYSTAL.lblproceso.Visible = False
  RCRYSTAL.ProgBar.Visible = False
  Pantalla.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
  
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
OJO:
If Err.Number = 70 Then
  MsgBox "Hoja de Calculo : " & wsfile1 & "  esta Abierta debe cerrar para Procesar Nuevamente ", 48, Pub_Titulo
  GoTo CANCELA
End If
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 xl.APPLICATION.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
'' Unload FrmImp2
Exit Sub
End Sub


