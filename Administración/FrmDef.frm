VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmSub 
   Caption         =   "Configuración de Transacciones"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "FrmDef.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   6660
   ScaleWidth      =   9510
   WindowState     =   2  'Maximized
   Begin VB.ListBox LisTransa 
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
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   1800
      TabIndex        =   70
      Tag             =   "0"
      Top             =   5760
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Frame MSTDEF 
      Caption         =   "Seleccione  Trnsacción :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   6750
      TabIndex        =   41
      Top             =   1200
      Visible         =   0   'False
      Width           =   4455
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   3375
         Left            =   105
         TabIndex        =   42
         Top             =   285
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   5953
         _Version        =   393216
         Rows            =   50
         Cols            =   4
         BackColor       =   -2147483624
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Consultar 
      Height          =   735
      Left            =   5520
      TabIndex        =   43
      Top             =   5880
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1296
      _Version        =   393216
      Cols            =   4
      FixedCols       =   3
   End
   Begin MSFlexGridLib.MSFlexGrid ListExiste 
      Height          =   495
      Left            =   840
      TabIndex        =   40
      Top             =   6120
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   400
      Left            =   5280
      TabIndex        =   32
      Top             =   5280
      Width           =   1500
   End
   Begin VB.Timer PARPADEA 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   240
      Top             =   5880
   End
   Begin VB.TextBox TxtDef 
      Height          =   285
      Index           =   39
      Left            =   3840
      TabIndex        =   38
      Text            =   " "
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Ce&rrar"
      Height          =   400
      Left            =   7440
      TabIndex        =   34
      Top             =   5280
      Width           =   1500
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   400
      Left            =   600
      TabIndex        =   30
      Top             =   5280
      Width           =   1500
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   400
      Left            =   3000
      TabIndex        =   31
      Top             =   5280
      Width           =   1500
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   0
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   840
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7435
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Definición de Transacción"
      TabPicture(0)   =   "FrmDef.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "checrecont"
      Tab(0).Control(1)=   "cmdtipo"
      Tab(0).Control(2)=   "txtsutgru"
      Tab(0).Control(3)=   "def_is"
      Tab(0).Control(4)=   "def_calidad"
      Tab(0).Control(5)=   "def_descto"
      Tab(0).Control(6)=   "def_signo_lot"
      Tab(0).Control(7)=   "def_signo_caja"
      Tab(0).Control(8)=   "def_signo_arm"
      Tab(0).Control(9)=   "def_signo_car"
      Tab(0).Control(10)=   "def_signo_ccm"
      Tab(0).Control(11)=   "def_abreviado"
      Tab(0).Control(12)=   "def_tipdoc"
      Tab(0).Control(13)=   "def_tipmov"
      Tab(0).Control(14)=   "def_cp"
      Tab(0).Control(15)=   "def_car_situacion"
      Tab(0).Control(16)=   "def_concepto"
      Tab(0).Control(17)=   "def_mortal"
      Tab(0).Control(18)=   "def_jabas"
      Tab(0).Control(19)=   "def_pollos"
      Tab(0).Control(20)=   "def_codpro"
      Tab(0).Control(21)=   "Frame1"
      Tab(0).Control(22)=   "Frame2(0)"
      Tab(0).Control(23)=   "Frame2(1)"
      Tab(0).Control(24)=   "def_art_gru"
      Tab(0).Control(25)=   "Label3"
      Tab(0).Control(26)=   "Label2"
      Tab(0).Control(27)=   "Label1(43)"
      Tab(0).Control(28)=   "Label1(39)"
      Tab(0).Control(29)=   "Label1(26)"
      Tab(0).Control(30)=   "Label1(30)"
      Tab(0).Control(31)=   "Label1(29)"
      Tab(0).Control(32)=   "Label1(28)"
      Tab(0).Control(33)=   "Label1(27)"
      Tab(0).Control(34)=   "Label1(7)"
      Tab(0).Control(35)=   "Label1(6)"
      Tab(0).Control(36)=   "Label1(5)"
      Tab(0).Control(37)=   "Label1(3)"
      Tab(0).Control(38)=   "Label1(4)"
      Tab(0).Control(39)=   "Label1(8)"
      Tab(0).Control(40)=   "Label1(31)"
      Tab(0).Control(41)=   "Label1(32)"
      Tab(0).Control(42)=   "Label1(34)"
      Tab(0).Control(43)=   "Label1(35)"
      Tab(0).Control(44)=   "Label1(36)"
      Tab(0).Control(45)=   "Label1(33)"
      Tab(0).Control(46)=   "Label1(37)"
      Tab(0).ControlCount=   47
      TabCaption(1)   =   "&Definición Contable x Compañia "
      TabPicture(1)   =   "FrmDef.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label4(8)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label4(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label4(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label4(2)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label4(3)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label4(4)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label4(5)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label4(6)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label4(7)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "fracnt"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "FRAOPT"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Definición de Series"
      TabPicture(2)   =   "FrmDef.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtDef(42)"
      Tab(2).Control(1)=   "TxtDef(41)"
      Tab(2).Control(2)=   "TxtDef(40)"
      Tab(2).Control(3)=   "Label1(42)"
      Tab(2).Control(4)=   "Label1(41)"
      Tab(2).Control(5)=   "Label1(40)"
      Tab(2).ControlCount=   6
      Begin VB.CheckBox checrecont 
         Caption         =   "Solo Ventas: Credito Contado"
         Height          =   255
         Left            =   -66480
         TabIndex        =   142
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Frame FRAOPT 
         Caption         =   "Sub Transacción :"
         Height          =   855
         Left            =   7800
         TabIndex        =   139
         Top             =   3240
         Width           =   3015
         Begin VB.OptionButton opt 
            Caption         =   "NO hace Asientos Contables"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   141
            Top             =   480
            Width           =   2415
         End
         Begin VB.OptionButton opt 
            Caption         =   "Requiere Asiento Contable"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   140
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.ComboBox cmdtipo 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   -66480
         Style           =   2  'Dropdown List
         TabIndex        =   136
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox txtsutgru 
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
         Left            =   -64920
         MaxLength       =   6
         TabIndex        =   134
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TxtDef 
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
         Index           =   42
         Left            =   -73440
         MaxLength       =   30
         TabIndex        =   132
         Text            =   "def_serie_recibos"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox TxtDef 
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
         Index           =   41
         Left            =   -73440
         MaxLength       =   30
         TabIndex        =   130
         Text            =   "def_serie_recibos"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox TxtDef 
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
         Index           =   40
         Left            =   -73440
         MaxLength       =   30
         TabIndex        =   128
         Text            =   "def_serie"
         Top             =   840
         Width           =   855
      End
      Begin VB.Frame fracnt 
         Caption         =   "Definir Asientos Contables"
         Height          =   3735
         Left            =   120
         TabIndex        =   71
         Top             =   360
         Width           =   7575
         Begin VB.TextBox TxtDef 
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
            Index           =   38
            Left            =   6360
            MaxLength       =   8
            TabIndex        =   107
            Text            =   "def_campo12"
            Top             =   3120
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   37
            Left            =   5760
            MaxLength       =   1
            TabIndex        =   106
            Text            =   "def_dh12"
            Top             =   3120
            Width           =   375
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   36
            Left            =   4560
            MaxLength       =   12
            TabIndex        =   105
            Text            =   "def_cta12"
            Top             =   3120
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   35
            Left            =   6360
            MaxLength       =   8
            TabIndex        =   104
            Text            =   "def_campo11"
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   34
            Left            =   5760
            MaxLength       =   1
            TabIndex        =   103
            Text            =   "def_dh11"
            Top             =   2640
            Width           =   375
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   33
            Left            =   4560
            MaxLength       =   12
            TabIndex        =   102
            Text            =   "def_cta11"
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   32
            Left            =   6360
            MaxLength       =   8
            TabIndex        =   101
            Text            =   "def_campo10"
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   31
            Left            =   5760
            MaxLength       =   1
            TabIndex        =   100
            Text            =   "def_dh10"
            Top             =   2160
            Width           =   375
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   30
            Left            =   4560
            MaxLength       =   12
            TabIndex        =   99
            Text            =   "def_cta10"
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   29
            Left            =   6360
            MaxLength       =   8
            TabIndex        =   98
            Text            =   "def_campo9"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   28
            Left            =   5760
            MaxLength       =   1
            TabIndex        =   97
            Text            =   "def_dh9"
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   27
            Left            =   4560
            MaxLength       =   12
            TabIndex        =   96
            Text            =   "def_cta9"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   26
            Left            =   6360
            MaxLength       =   8
            TabIndex        =   95
            Text            =   "def_campo8"
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   25
            Left            =   5760
            MaxLength       =   1
            TabIndex        =   94
            Text            =   "def_dh8"
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   24
            Left            =   4560
            MaxLength       =   12
            TabIndex        =   93
            Text            =   "def_cta8"
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   23
            Left            =   6360
            MaxLength       =   8
            TabIndex        =   92
            Text            =   "def_campo7"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   22
            Left            =   5760
            MaxLength       =   1
            TabIndex        =   91
            Text            =   "def_dh7"
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   21
            Left            =   4560
            MaxLength       =   12
            TabIndex        =   90
            Text            =   "def_cta7"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   20
            Left            =   2520
            MaxLength       =   8
            TabIndex        =   89
            Text            =   "def_campo6"
            Top             =   3120
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   19
            Left            =   1920
            MaxLength       =   1
            TabIndex        =   88
            Text            =   "def_dh6"
            Top             =   3120
            Width           =   375
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   18
            Left            =   720
            MaxLength       =   12
            TabIndex        =   87
            Text            =   "def_cta6"
            Top             =   3120
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   17
            Left            =   2520
            MaxLength       =   8
            TabIndex        =   86
            Text            =   "def_campo5"
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   16
            Left            =   1920
            MaxLength       =   1
            TabIndex        =   85
            Text            =   "def_dh5"
            Top             =   2640
            Width           =   375
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   15
            Left            =   720
            MaxLength       =   12
            TabIndex        =   84
            Text            =   "def_cta5"
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   14
            Left            =   2520
            MaxLength       =   8
            TabIndex        =   83
            Text            =   "def_campo4"
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   13
            Left            =   1920
            MaxLength       =   1
            TabIndex        =   82
            Text            =   "def_dh4"
            Top             =   2160
            Width           =   375
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   12
            Left            =   720
            MaxLength       =   12
            TabIndex        =   81
            Text            =   "def_cta4"
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   11
            Left            =   2520
            MaxLength       =   8
            TabIndex        =   80
            Text            =   "def_campo3"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   10
            Left            =   1920
            MaxLength       =   1
            TabIndex        =   79
            Text            =   "def_dh3"
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   9
            Left            =   720
            MaxLength       =   12
            TabIndex        =   78
            Text            =   "def_cta3"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   8
            Left            =   2520
            MaxLength       =   8
            TabIndex        =   77
            Text            =   "def_campo2"
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   7
            Left            =   1920
            MaxLength       =   1
            TabIndex        =   76
            Text            =   "def_dh2"
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   6
            Left            =   720
            MaxLength       =   12
            TabIndex        =   75
            Text            =   "def_cta2"
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   5
            Left            =   2520
            MaxLength       =   8
            TabIndex        =   74
            Text            =   "def_campo1"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   4
            Left            =   1920
            MaxLength       =   1
            TabIndex        =   73
            Text            =   "def_dh1"
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox TxtDef 
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
            Index           =   3
            Left            =   720
            MaxLength       =   12
            TabIndex        =   72
            Text            =   "def_cta1"
            Top             =   720
            Width           =   975
         End
         Begin ComctlLib.ProgressBar bp 
            Height          =   195
            Left            =   2160
            TabIndex        =   126
            Top             =   3480
            Visible         =   0   'False
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   344
            _Version        =   327682
            Appearance      =   0
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Campo"
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
            Index           =   9
            Left            =   2760
            TabIndex        =   125
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Debe. Haber. "
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
            Height          =   390
            Index           =   10
            Left            =   1920
            TabIndex        =   124
            Top             =   240
            Width           =   645
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cta. Contab."
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
            Height          =   390
            Index           =   11
            Left            =   960
            TabIndex        =   123
            Top             =   240
            Width           =   705
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Campo"
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
            Index           =   12
            Left            =   6720
            TabIndex        =   122
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Debe. Haber. "
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
            Height          =   390
            Index           =   13
            Left            =   5880
            TabIndex        =   121
            Top             =   240
            Width           =   645
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cta. Contab."
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
            Height          =   390
            Index           =   14
            Left            =   4800
            TabIndex        =   120
            Top             =   240
            Width           =   705
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N. 07.-"
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
            Index           =   15
            Left            =   3960
            TabIndex        =   119
            Top             =   720
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N. 08.-"
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
            Index           =   16
            Left            =   3960
            TabIndex        =   118
            Top             =   1200
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N. 09.-"
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
            Index           =   17
            Left            =   3960
            TabIndex        =   117
            Top             =   1680
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N. 10.-"
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
            Index           =   18
            Left            =   3960
            TabIndex        =   116
            Top             =   2160
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N. 11.-"
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
            Index           =   19
            Left            =   3960
            TabIndex        =   115
            Top             =   2640
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N. 12.-"
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
            Index           =   38
            Left            =   3960
            TabIndex        =   114
            Top             =   3120
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N. 01.-"
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
            Index           =   20
            Left            =   120
            TabIndex        =   113
            Top             =   720
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N. 02.-"
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
            Index           =   21
            Left            =   120
            TabIndex        =   112
            Top             =   1200
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N. 03.-"
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
            Index           =   22
            Left            =   120
            TabIndex        =   111
            Top             =   1680
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N. 04.-"
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
            Index           =   23
            Left            =   120
            TabIndex        =   110
            Top             =   2160
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N. 05.-"
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
            Index           =   24
            Left            =   120
            TabIndex        =   109
            Top             =   2640
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N. 06.-"
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
            Index           =   25
            Left            =   120
            TabIndex        =   108
            Top             =   3120
            Width           =   600
         End
      End
      Begin VB.TextBox def_is 
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
         Left            =   -67080
         MaxLength       =   1
         TabIndex        =   69
         Text            =   "def_is"
         Top             =   1080
         Width           =   975
      End
      Begin VB.ComboBox def_calidad 
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
         Left            =   -67200
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   1440
         Width           =   1185
      End
      Begin VB.TextBox def_descto 
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
         Left            =   -67080
         MaxLength       =   8
         TabIndex        =   65
         Text            =   "def_descto"
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox def_signo_lot 
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
         Left            =   -71040
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2520
         Width           =   700
      End
      Begin VB.ComboBox def_signo_caja 
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
         Left            =   -71040
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2040
         Width           =   700
      End
      Begin VB.ComboBox def_signo_arm 
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
         Left            =   -71040
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1560
         Width           =   700
      End
      Begin VB.ComboBox def_signo_car 
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
         Left            =   -71040
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1080
         Width           =   700
      End
      Begin VB.ComboBox def_signo_ccm 
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
         Left            =   -71040
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   700
      End
      Begin VB.ComboBox def_abreviado 
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
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2040
         Width           =   1815
      End
      Begin VB.ComboBox def_tipdoc 
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
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1560
         Width           =   1815
      End
      Begin VB.ComboBox def_tipmov 
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
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox def_cp 
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
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox def_car_situacion 
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
         Left            =   -69240
         MaxLength       =   6
         TabIndex        =   14
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox def_concepto 
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
         Left            =   -69360
         MaxLength       =   10
         TabIndex        =   18
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox def_mortal 
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
         Left            =   -69240
         MaxLength       =   6
         TabIndex        =   15
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox def_jabas 
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
         Left            =   -69240
         MaxLength       =   10
         TabIndex        =   16
         Text            =   "def_jabas"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox def_pollos 
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
         Left            =   -69240
         MaxLength       =   10
         TabIndex        =   17
         Text            =   "def_pollos"
         Top             =   2040
         Width           =   615
      End
      Begin VB.ComboBox def_codpro 
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
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2470
         Width           =   1815
      End
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   -74880
         TabIndex        =   47
         Top             =   3240
         Width           =   4575
         Begin VB.CommandButton cmdCopiar 
            Caption         =   "Copiar Datos "
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
            Height          =   350
            Left            =   1560
            TabIndex        =   28
            Top             =   210
            Width           =   1455
         End
         Begin VB.CheckBox chetransa 
            Caption         =   "Incluir Nueva Transación"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   160
            Width           =   1455
         End
         Begin VB.CommandButton cmdconsultar 
            Caption         =   "Ver Trans."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   3240
            TabIndex        =   29
            Top             =   210
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Precios"
         Height          =   1215
         Index           =   0
         Left            =   -70200
         TabIndex        =   46
         Top             =   2880
         Width           =   2055
         Begin VB.OptionButton def_precio 
            Caption         =   "Fijo."
            Height          =   255
            Index           =   4
            Left            =   720
            TabIndex        =   127
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton def_precio 
            Caption         =   "Por Articulo"
            Height          =   255
            Index           =   3
            Left            =   720
            TabIndex        =   24
            Top             =   120
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton def_precio 
            Caption         =   "Digitado"
            Height          =   255
            Index           =   2
            Left            =   720
            TabIndex        =   25
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton def_precio 
            Caption         =   "A Costo"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   26
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "P r e c i o s"
         Height          =   1575
         Index           =   1
         Left            =   -68040
         TabIndex        =   45
         Top             =   2520
         Width           =   1335
         Begin VB.CheckBox Check1 
            Caption         =   "Precio 1"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Precio 2"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   480
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Precio 3"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Precio 4"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   22
            Top             =   960
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Precio 5"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   23
            Top             =   1200
            Width           =   1095
         End
      End
      Begin VB.ComboBox def_art_gru 
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
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "53 = Flete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   195
         Index           =   7
         Left            =   8070
         TabIndex        =   150
         Top             =   2595
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "30 = Impto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   195
         Index           =   6
         Left            =   8070
         TabIndex        =   149
         Top             =   2325
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "29 = Gastos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   195
         Index           =   5
         Left            =   8070
         TabIndex        =   148
         Top             =   2070
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "28 = Bruto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   195
         Index           =   4
         Left            =   8070
         TabIndex        =   147
         Top             =   1800
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "27 = Neto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   195
         Index           =   3
         Left            =   8070
         TabIndex        =   146
         Top             =   1545
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  8 = Importe"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   195
         Index           =   2
         Left            =   8070
         TabIndex        =   145
         Top             =   1275
         Width           =   1290
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  7 = Importe_Amort"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   195
         Index           =   1
         Left            =   8070
         TabIndex        =   144
         Top             =   1020
         Width           =   2010
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "CAMPO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0057C457&
         Height          =   240
         Index           =   0
         Left            =   8880
         TabIndex        =   143
         Top             =   780
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "Relación con Tipo de Descto."
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   -66480
         TabIndex        =   138
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Solo Para Ventas."
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -66480
         TabIndex        =   137
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FAMILIA:"
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
         Index           =   43
         Left            =   -65880
         TabIndex        =   135
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Serie por CAja :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Index           =   42
         Left            =   -74640
         TabIndex        =   133
         Top             =   1680
         Width           =   1095
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Serie por Cobranzas :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   630
         Index           =   41
         Left            =   -74640
         TabIndex        =   131
         Top             =   1200
         Width           =   1050
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Serie por Mercaderia :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Index           =   40
         Left            =   -74640
         TabIndex        =   129
         Top             =   720
         Width           =   1035
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "I/S"
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
         Index           =   39
         Left            =   -68040
         TabIndex        =   67
         Top             =   1080
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Calidad"
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
         Index           =   26
         Left            =   -68040
         TabIndex        =   66
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lotes   :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   30
         Left            =   -71880
         TabIndex        =   64
         Top             =   2520
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Caja       :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   29
         Left            =   -71880
         TabIndex        =   63
         Top             =   2040
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Articulo :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   28
         Left            =   -71880
         TabIndex        =   62
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cartera  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   27
         Left            =   -71880
         TabIndex        =   61
         Top             =   1080
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Abreviado :"
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
         Index           =   7
         Left            =   -74880
         TabIndex        =   60
         Top             =   2040
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Doc.  :"
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
         Index           =   6
         Left            =   -74880
         TabIndex        =   59
         Top             =   1560
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Movi. :"
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
         Index           =   5
         Left            =   -74880
         TabIndex        =   58
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Clie./Prov. :"
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
         Index           =   3
         Left            =   -74880
         TabIndex        =   57
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cta. Cte. : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   4
         Left            =   -71880
         TabIndex        =   56
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Situacion "
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
         Index           =   8
         Left            =   -70200
         TabIndex        =   55
         Top             =   600
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
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
         Index           =   31
         Left            =   -70200
         TabIndex        =   54
         Top             =   2520
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Compra"
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
         Height          =   390
         Index           =   32
         Left            =   -70200
         TabIndex        =   53
         Top             =   1080
         Width           =   1005
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desct. Facturación"
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
         Height          =   435
         Index           =   34
         Left            =   -70320
         TabIndex        =   52
         Top             =   2040
         Width           =   1155
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Jabas"
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
         Index           =   35
         Left            =   -70200
         TabIndex        =   51
         Top             =   1560
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proceso   :"
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
         Index           =   36
         Left            =   -74880
         TabIndex        =   50
         Top             =   2520
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "P=3A  , 2da.P es 100% "
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
         Height          =   585
         Index           =   33
         Left            =   -68280
         TabIndex        =   49
         Top             =   480
         Width           =   1275
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   37
         Left            =   -74880
         TabIndex        =   48
         Top             =   2880
         Width           =   660
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         Height          =   2325
         Index           =   8
         Left            =   7905
         TabIndex        =   151
         Top             =   690
         Width           =   2940
      End
   End
   Begin VB.Frame F1 
      Height          =   735
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   10935
      Begin VB.TextBox TxtDef 
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
         Index           =   2
         Left            =   3360
         MaxLength       =   30
         TabIndex        =   2
         Text            =   "def_descripcion"
         Top             =   360
         Width           =   6015
      End
      Begin VB.TextBox TxtDef 
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
         Index           =   1
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "def_secuencia"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox TxtDef 
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
         Index           =   0
         Left            =   360
         MaxLength       =   4
         TabIndex        =   0
         Text            =   "Def_codtra"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción "
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
         Index           =   2
         Left            =   3240
         TabIndex        =   37
         Top             =   120
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Secuencia  :"
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
         Index           =   1
         Left            =   1800
         TabIndex        =   36
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Trans.  :"
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
         Index           =   0
         Left            =   240
         TabIndex        =   35
         Top             =   120
         Width           =   1185
      End
   End
   Begin VB.Label LblMensaje 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   39
      Top             =   5040
      Width           =   4635
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IND As Integer
Dim CU As Integer
Dim llave1
Dim TAB_CHECK(5) As String * 1

Public Sub LLENADOS(cont As ComboBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!TAB_NOMLARGO & String(15, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub

Public Function PROC_DEF()
Dim tempo As String
Dim NUMCAMPO As Integer
Dim wBusca As String
Dim OJO As String * 1
Dim WW As String * 1

If PUB_CODTRA = 0 Then
   wBusca = "SELECT * FROM SUB_TRANSA  ORDER BY SUT_SECUENCIA "
Else
   wBusca = "SELECT * FROM SUB_TRANSA WHERE SUT_CODTRA = ?  ORDER BY SUT_CODTRA " ' SUB_SECUENCIA, SUB_DESCRIPCION "
End If
If UNICO <> wBusca Then
   Set PSX = CN.CreateQuery("", wBusca)
   If PUB_CODTRA <> 0 Then PSX.rdoParameters(0) = 0
End If
If PUB_CODTRA <> 0 Then
  PSX.rdoParameters(0) = PUB_CODTRA
End If
If UNICO <> wBusca Then
   Set X = PSX.OpenResultset(rdOpenKeyset)
End If
If UNICO = wBusca Then
   X.Requery
   If X.RowCount > 0 Then
      X.MoveFirst
   End If
End If
If X.EOF = True Then
   Grid1.Clear
'   Grid1.Visible = True
'   Grid1.Rows = 2
'   Grid1.TextMatrix(1, 3) = "No hay registros"
'   Screen.MousePointer = 0
'   llave1 = ""
'   UNICO = wBusca
   PUB_CODTRA = -1
   Exit Function
End If

If X.rdoColumns(0) = llave1 Then
    Exit Function
End If
If X.RowCount > 0 Then
   llave1 = X.rdoColumns(0)
End If
UNICO = wBusca
Grid1.Rows = 2
fila = 0
Grid1.Clear
Grid1.TextMatrix(fila, 1) = "Tras."
Grid1.TextMatrix(fila, 2) = "Sec"
Grid1.TextMatrix(fila, 3) = "Descripción"
Grid1.Rows = 2
Grid1.Cols = 4
Grid1.ColWidth(0) = 250
Grid1.ColWidth(1) = 700
Grid1.ColWidth(2) = 300
Grid1.ColWidth(3) = 2600
fila = 0
Grid1.Visible = False
Do Until X.EOF 'Or fila = 50
    fila = fila + 1
    Grid1.TextMatrix(fila, 0) = fila 'Nulo_Valors(X.rdoColumns(3))
    If Nulo_Valors(X!SUT_FLAG_CONTABLE) = "A" Then
       WW = "A"
    Else
       Grid1.Col = 1
       Grid1.Row = fila
       Grid1.CellBackColor = QBColor(7)
       WW = "X"
    End If
    Grid1.TextMatrix(fila, 1) = X.rdoColumns(0) & "-" & WW
    Grid1.TextMatrix(fila, 2) = Trim(X.rdoColumns(1))
    Grid1.TextMatrix(fila, 3) = X.rdoColumns(2)
    X.MoveNext
    Grid1.Rows = Grid1.Rows + 1
Loop
Grid1.Visible = True
Grid1.TextMatrix(fila + 1, 3) = "                * * *    END    * * "
End Function

Public Sub GRABAR_DEF()
If Left(CmdModificar.Caption, 2) = "&G" Then
   SUT_LLAVE.Edit
   If cnt_llave.EOF Then
    cnt_llave.AddNew
    cnt_llave!CNT_CODCIA = LK_CODCIA
    cnt_llave!cnt_CODTRA = Val(FrmSub.TxtDef(0).text)
    cnt_llave!cnt_SECUENCIA = Val(FrmSub.TxtDef(1).text)
    cnt_llave.Update
    SQ_OPER = 1
    PUB_CODTRA = Val(FrmSub.TxtDef(0).text)
    PUB_SECUENCIA = Val(FrmSub.TxtDef(1).text)
    PUB_CODCIA = LK_CODCIA
    LEER_CNT_LLAVE
   End If
    cnt_llave.Edit
    
Else
   SUT_LLAVE.AddNew
   cnt_llave.AddNew
End If
cnt_llave!CNT_CODCIA = LK_CODCIA
cnt_llave!cnt_CODTRA = Val(FrmSub.TxtDef(0).text)
cnt_llave!cnt_SECUENCIA = Val(FrmSub.TxtDef(1).text)

For i = 3 To 38
   If cnt_llave(i).Type = 1 Then
      If opt(0).Value Then
       cnt_llave(i) = Nulo_Valors(TxtDef(i).text)
      Else
        cnt_llave(i) = " "
      End If
   Else
      If opt(0).Value Then
       cnt_llave(i) = Val(TxtDef(i).text)
      Else
        cnt_llave(i) = 0
      End If
   End If
Next i
cnt_llave!cnt_serie = Val(TxtDef(40).text)
cnt_llave!cnt_serie_recibo = Val(TxtDef(41).text)
cnt_llave!cnt_serie_caja = Val(TxtDef(42).text)
cnt_llave.Update
SUT_LLAVE!SUT_DESCRIPCION = Trim(TxtDef(2).text)
SUT_LLAVE!SUT_cp = Left(def_cp.text, 1)
SUT_LLAVE!SUT_tipdoc = Left(def_tipdoc.text, 2)
SUT_LLAVE!SUT_abreviado = Trim(def_abreviado.text)
SUT_LLAVE!SUT_signo_ccm = Val(def_signo_ccm.text)
SUT_LLAVE!SUT_signo_arm = Val(def_signo_arm.text)
SUT_LLAVE!SUT_signo_car = Val(def_signo_car.text)
SUT_LLAVE!SUT_signo_caja = Val(def_signo_caja.text)
SUT_LLAVE!SUT_tipmov = Val(Right(def_tipmov.text, 4))
SUT_LLAVE!SUT_tipmov_REF = Val(Right(cmdtipo.text, 4))
SUT_LLAVE!SUT_CALIDAD = Val(Right(def_calidad, 4))

SUT_LLAVE!sut_descto = Val(def_descto.text)
SUT_LLAVE!sut_is = Val(def_is.text)
SUT_LLAVE!SUT_codpro = Val(Right(def_codpro.text, 4))
SUT_LLAVE!sut_art_gru = Val(Right(def_art_gru.text, 4))
SUT_LLAVE!SUT_FLAG_TIPO = Trim(def_mortal.text)
'SUT_LLAVE!SUT_SERIE = Val(FrmSub.TxtDef(40).text)
SUT_LLAVE!SUT_FLAG_CC = checrecont.Value

If def_precio(2).Value Then
 SUT_LLAVE!SUT_precio = "2"
ElseIf def_precio(1).Value Then
 SUT_LLAVE!SUT_precio = "1"
ElseIf def_precio(3).Value Then
 SUT_LLAVE!SUT_precio = "3"
Else
 SUT_LLAVE!SUT_precio = "4"
End If
TAB_CHECK(0) = "0"
TAB_CHECK(1) = "0"
TAB_CHECK(2) = "0"
TAB_CHECK(3) = "0"
TAB_CHECK(4) = "0"
If Check1(0).Value = 1 Then TAB_CHECK(0) = "1"
If Check1(1).Value = 1 Then TAB_CHECK(1) = "1"
If Check1(2).Value = 1 Then TAB_CHECK(2) = "1"
If Check1(3).Value = 1 Then TAB_CHECK(3) = "1"
If Check1(4).Value = 1 Then TAB_CHECK(4) = "1"
   
SUT_LLAVE!SUT_PRECIOS = TAB_CHECK(0) & TAB_CHECK(1) & TAB_CHECK(2) & TAB_CHECK(3) & TAB_CHECK(4)
SUT_LLAVE!SUT_jabas = def_jabas.text
SUT_LLAVE!SUT_DESFAC = def_pollos.text
SUT_LLAVE!sut_descto = def_descto.text
SUT_LLAVE!sut_is = def_is.text
SUT_LLAVE!sut_art_gru = Val(txtsutgru.text)

If opt(0).Value = True Then
  SUT_LLAVE!SUT_FLAG_CONTABLE = "A"
Else
  SUT_LLAVE!SUT_FLAG_CONTABLE = " "
End If


SUT_LLAVE.Update

End Sub


Private Sub cmdcancelar_Click()
If Left(CmdModificar.Caption, 2) = "&M" Then
    LIMPIA_DEF
    TxtDef(0).Locked = False
    BLOQUEA_TEXT
    MENSAJE_DEF "Inicializar ... !!!    "
    TxtDef(0).Enabled = True
    TxtDef(0).SetFocus
    FrmSub.SSTab1.Tab = 0
    cmdCopiar.Enabled = False
    chetransa.Enabled = False
    cmdconsultar.Enabled = False
    Exit Sub
End If
     Screen.MousePointer = 11
     If Left(CmdModificar.Caption, 2) = "&G" Then
        CmdModificar.Caption = "&Modificar"
        LLENA_DEF 1
        cmdCopiar.Enabled = True
        chetransa.Enabled = True
        cmdconsultar.Enabled = True
        TxtDef(0).Locked = True
     End If
     cmdEliminar.Enabled = True
     CmdModificar.Enabled = True
     BLOQUEA_TEXT
     FrmSub.SSTab1.Tab = 0
     MENSAJE_DEF "Inicializar... !!! "
    Screen.MousePointer = 0

End Sub

Private Sub cmdcerrar_Click()
Unload FrmSub
End Sub

Private Sub cmdCerrar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    TxtDef(0).SetFocus
End If

End Sub

Private Sub cmdConsultar_Click()
Dim i As Integer
If Trim(TxtDef(0).text) = "" Then
  Screen.MousePointer = 0
  MsgBox " Digitar Codigo..", 48, Pub_Titulo
  TxtDef(0).SetFocus
  Exit Sub
End If
Screen.MousePointer = 11
Consultar.Width = 9300 '9495
Consultar.Height = 2775
Consultar.Left = 10
Consultar.Top = 1440
Consultar.RowHeight(0) = 400
SQ_OPER = 2
PUB_CODTRA = Val(TxtDef(0).text)
LEER_SUT_LLAVE
If SUT_MAYOR.EOF Then
 Screen.MousePointer = 0
 Exit Sub
End If
fila = 1
Consultar.Rows = 1
Consultar.Cols = SUT_MAYOR.rdoColumns.Count
Do Until SUT_MAYOR.EOF
 Consultar.Rows = Consultar.Rows + 1
 For i = 0 To SUT_MAYOR.rdoColumns.Count - 1
  Consultar.TextMatrix(0, i) = Mid(SUT_MAYOR.rdoColumns(i).Name, 5, Len(SUT_MAYOR.rdoColumns(i).Name))
  If Not IsNull(SUT_MAYOR.rdoColumns(i)) Then
    If IsNumeric(SUT_MAYOR.rdoColumns(i)) Then
     Consultar.TextMatrix(fila, i) = Val(SUT_MAYOR.rdoColumns(i))
    Else
     Consultar.TextMatrix(fila, i) = SUT_MAYOR.rdoColumns(i)
    End If
  End If
  Consultar.ColWidth(i) = 700
 Next i
 fila = fila + 1
 SUT_MAYOR.MoveNext
Loop
Consultar.ColWidth(0) = 500
Consultar.ColWidth(1) = 400
Consultar.ColWidth(2) = 1500

Screen.MousePointer = 0
Consultar.Visible = True
Consultar.Col = 3
Consultar.Row = 1
Consultar.SetFocus

End Sub

Private Sub cmdCopiar_Click()
Dim Mensaje, Título, valorpred, mivalor
Dim otro_llave As rdoResultset
Dim msg, estilo, respuesta
Dim wcodtra As Integer
wcodtra = 0
If SUT_LLAVE.EOF Then
 MsgBox "Intente nuevamente ... ", 48, Pub_Titulo
 cmdcancelar_Click
 Exit Sub
End If
If chetransa.Value = 1 Then
    Mensaje = "Introduzca la Nueva Transacción , debe ser numerico de 4 Digitos : "
    Título = "Definición de Transacción "
    valorpred = " "
    mivalor = InputBox(Mensaje, Título, valorpred)
    If mivalor = "" Then
       Exit Sub
    End If
    If Len(Trim(mivalor)) > 4 Or Len(Trim(mivalor)) < 4 Or Len(Trim(mivalor)) = 0 Then
       MsgBox "Son solo 4 digitos .... Reintente hacer la copia !!!", 48, Pub_Titulo
       Exit Sub
    End If
    If InStr(1, mivalor, ".") > 0 Then
       MsgBox "Debe ser Entero  .... Reintente hacer la copia !!!", 48, Pub_Titulo
       Exit Sub
    End If
    If Not IsNumeric(mivalor) Then
       MsgBox "Debe ser Numerico .... Reintente hacer la copia !!!", 48, Pub_Titulo
       Exit Sub
    End If
    PUB_CODTRA = Trim(mivalor)
    wcodtra = 1
    msg = "Ahora su "
Else
    PUB_CODTRA = Trim(TxtDef(0).text)
    msg = "Ingrese la "
End If
 Mensaje = msg & "Secuencia para la Transacción , debe ser numerico  : "
 Título = "Copia - Definición de Transacción "
 valorpred = " "
 mivalor = InputBox(Mensaje, Título, valorpred)
 If mivalor = "" Then
    Exit Sub
 End If

 If Not IsNumeric(mivalor) Then
    MsgBox " No Procede . . ., Intente nuevamente ", 48, Pub_Titulo
    Exit Sub
 End If
 If Val(mivalor) < 0 Then
    MsgBox "Solo valores Positivos  .. Reintente hacer la copia !!!", 48, Pub_Titulo
    Exit Sub
 End If
 If InStr(1, mivalor, ".") > 0 Then
    MsgBox "Debe ser Entero  .... Reintente hacer la copia !!!", 48, Pub_Titulo
    Exit Sub
 End If
Screen.MousePointer = 11

pub_cadena = "SELECT * FROM sub_Transa"
Set otro_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
otro_llave.Requery
 
otro_llave.AddNew
For fila = 0 To SUT_LLAVE.rdoColumns.Count - 1
 If fila = 1 Then
   otro_llave.rdoColumns(fila) = Val(mivalor)
 ElseIf fila = 0 And wcodtra = 1 Then
   otro_llave.rdoColumns(fila) = PUB_CODTRA
 Else
 'xxxxxxxxxxxxxxxx
  If IsNull(SUT_LLAVE.rdoColumns(fila)) Then
  Else
    otro_llave.rdoColumns(fila) = SUT_LLAVE.rdoColumns(fila)
  End If
 End If
Next
 
 SQ_OPER = 1
 PUB_SECUENCIA = Val(mivalor)
 LEER_SUT_LLAVE
 If Not SUT_LLAVE.EOF Then
    Screen.MousePointer = 0
    MsgBox "Secuencia de Transacción ya EXISTE... Reintente la copia !!!", 48, Pub_Titulo
    otro_llave.CancelUpdate
    cmdcancelar_Click
    Exit Sub
 End If
Screen.MousePointer = 11
otro_llave.Update
cnt_llave.AddNew
cnt_llave!CNT_CODCIA = LK_CODCIA
cnt_llave!cnt_CODTRA = PUB_CODTRA
cnt_llave!cnt_SECUENCIA = Val(mivalor)
For i = 3 To 38
   If cnt_llave(i).Type = 1 Then
      cnt_llave(i) = Nulo_Valors(TxtDef(i).text)
   Else
      cnt_llave(i) = Val(TxtDef(i).text)
   End If
Next i
cnt_llave.Update



Screen.MousePointer = 0
MsgBox "Copia Efectuada . Transacción  : " & PUB_CODTRA & "  Secuencia :" & mivalor, 48, Pub_Titulo
cmdcancelar_Click
otro_llave.Close
Exit Sub
End Sub

Private Sub cmdEliminar_Click()
If Len(TxtDef(0).text) = 0 Or Len(TxtDef(0).text) = 0 Then
'   MENSAJE_CLI "NO a seleccionado NADA ... !"
   Exit Sub
End If
  pub_mensaje = " ¿Desea Eliminar el Registro... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbYes Then   ' El usuario eligió
    Screen.MousePointer = 11
    SUT_LLAVE.Delete
    If Not cnt_llave.EOF Then cnt_llave.Delete
    FrmSub.TxtDef(0).text = ""
    FrmSub.TxtDef(0).Locked = False
    LIMPIA_DEF
    MENSAJE_DEF "Registro, ELIMINADO ... "
    Screen.MousePointer = 0
    FrmSub.TxtDef(0).SetFocus
   Exit Sub
  End If
  Screen.MousePointer = 0


End Sub

Private Sub cmdEliminar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    TxtDef(0).SetFocus
End If

End Sub

Private Sub CmdModificar_Click()
If Len(TxtDef(0).text) = 0 Or Len(TxtDef(0).text) = 0 Then
   MENSAJE_DEF "NO a seleccionado.. !"
   Exit Sub
End If
If Len(TxtDef(1).text) = 0 Or Len(TxtDef(1).text) = 0 Then
   MENSAJE_DEF "NO a seleccionado.. !"
   Exit Sub
End If

If Left(CmdModificar.Caption, 2) = "&M" Then
    CmdModificar.Caption = "&Grabar"
    cmdEliminar.Enabled = False
    cmdCancelar.Enabled = True
    DESBLOQUEA_TEXT
    TxtDef(0).Locked = True
    TxtDef(1).Enabled = False
    cmdCopiar.Enabled = False
    chetransa.Enabled = False
    cmdconsultar.Enabled = False
    FrmSub.TxtDef(2).SetFocus
Else
    '*Grabar las modificaciones
    If Not CONSIS Then
       Exit Sub
    End If

    Screen.MousePointer = 11
    GRABAR_DEF
    MENSAJE_DEF "Registro,MODIFICADO... "
    CmdModificar.Caption = "&Modificar"
    cmdCancelar.Enabled = True
    cmdEliminar.Enabled = True
    BLOQUEA_TEXT
    LIMPIA_DEF
    cmdCancelar.Enabled = True
    TxtDef(0).Locked = False
    TxtDef(0).Enabled = True
    TxtDef(0).SetFocus
    Screen.MousePointer = 0
    
End If


End Sub

Private Sub cmdModificar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    TxtDef(0).SetFocus
End If
End Sub

Private Sub Consultar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 Consultar.Visible = False
 TxtDef(0).SetFocus
End If
End Sub

Private Sub Consultar_LostFocus()
 Consultar.Visible = False
 TxtDef(0).SetFocus
End Sub

Private Sub chetransa_Click()
cmdCopiar.SetFocus
End Sub

Private Sub def_abreviado_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos
If KeyCode <> 45 Then
 Exit Sub
End If
wpos = def_abreviado.ListIndex
PUB_TIPREG = Mid(def_abreviado.ToolTipText, 13, Len(def_abreviado.ToolTipText))
Load FrmDatArti
PUB_CODCIA = "00"
FrmDatArti.Caption = "Def-Abreviado.  -  TAB_TIPREG = " & PUB_TIPREG
Load FrmDatArti
FrmDatArti.Show 1
DoEvents
PUB_CODCIA = "00"
LLENADOS FrmSub.def_abreviado, 13
def_abreviado.SetFocus
SendKeys "%{up}"

End Sub

Private Sub def_art_gru_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos
If KeyCode <> 45 Then
 Exit Sub
End If
wpos = def_art_gru.ListIndex
PUB_TIPREG = Mid(def_art_gru.ToolTipText, 13, Len(def_art_gru.ToolTipText))
Load FrmDatArti
PUB_CODCIA = LK_CODCIA
FrmDatArti.Caption = "Def_art_gru.  -  TAB_TIPREG = " & PUB_TIPREG
Load FrmDatArti
FrmDatArti.Show 1
DoEvents
PUB_CODCIA = LK_CODCIA
LLENADOS FrmSub.def_art_gru, 122
def_art_gru.SetFocus
SendKeys "%{up}"

End Sub

Private Sub def_car_situacion_KeyPress(KeyAscii As Integer)
Dim car As String, Longt As Integer
car = Chr$(KeyAscii)
car = UCase$(Chr$(KeyAscii))
KeyAscii = Asc(car)
If Not car < "0" And car > "9" Then
      If KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
        Beep
        Exit Sub
       End If
End If
End Sub

Private Sub def_codpro_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 45 Then
 Exit Sub
End If
PUB_TIPREG = Mid(def_codpro.ToolTipText, 13, Len(def_codpro.ToolTipText))
Load FrmDatArti
PUB_CODCIA = LK_CODCIA
FrmDatArti.Caption = "TIP-DOC.  -  TAB_TIPREG = " & PUB_TIPREG
Load FrmDatArti
FrmDatArti.Show 1
DoEvents
PUB_CODCIA = LK_CODCIA
LLENADOS FrmSub.def_codpro, 888

def_codpro.SetFocus
SendKeys "%{up}"
End Sub

Private Sub def_descto_KeyPress(KeyAscii As Integer)
Dim car As String, Longt As Integer
car = Chr$(KeyAscii)
car = UCase$(Chr$(KeyAscii))
KeyAscii = Asc(car)
If Not car < "0" And car > "9" And car <> "A" And car <> "P" And car <> "M" And car <> "F" Then
      If KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
        Beep
        Exit Sub
       End If
End If

End Sub

Private Sub def_precio_Click(Index As Integer)
If Index = 3 Then
   Frame2(1).Visible = True
Else
   Frame2(1).Visible = False
End If

End Sub

Private Sub def_tipdoc_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos
If KeyCode <> 45 Then
 Exit Sub
End If
wpos = def_tipdoc.ListIndex
PUB_TIPREG = Mid(def_tipdoc.ToolTipText, 13, Len(def_tipdoc.ToolTipText))
Load FrmDatArti
PUB_CODCIA = "00"
FrmDatArti.Caption = "TIP-DOC.  -  TAB_TIPREG = " & PUB_TIPREG
Load FrmDatArti
FrmDatArti.Show 1
DoEvents
PUB_CODCIA = "00"
LLENADOS FrmSub.def_tipdoc, 8
def_tipdoc.SetFocus
SendKeys "%{up}"

End Sub

Private Sub def_tipmov_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos
If KeyCode <> 45 Then
 Exit Sub
End If
wpos = def_tipmov.ListIndex
PUB_TIPREG = Mid(def_tipmov.ToolTipText, 13, Len(def_tipmov.ToolTipText))
Load FrmDatArti
PUB_CODCIA = "00"
FrmDatArti.Caption = "TIP-MOVIMIENTO -  TAB_TIPREG = " & PUB_TIPREG
Load FrmDatArti
FrmDatArti.Show 1
DoEvents
PUB_CODCIA = "00"
LLENADOS FrmSub.def_tipmov, 4
def_tipmov.SetFocus
SendKeys "%{up}"

End Sub

Private Sub def_calidad_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos
If KeyCode <> 45 Then
 Exit Sub
End If
wpos = def_calidad.ListIndex
PUB_TIPREG = Mid(def_calidad.ToolTipText, 13, Len(def_calidad.ToolTipText))
Load FrmDatArti
PUB_CODCIA = LK_CODCIA
FrmDatArti.Caption = "TIP-MOVIMIENTO -  TAB_TIPREG = " & PUB_TIPREG
Load FrmDatArti
FrmDatArti.Show 1
DoEvents
PUB_CODCIA = LK_CODCIA
LLENADOS FrmSub.def_calidad, 2
def_calidad.SetFocus
SendKeys "%{up}"

End Sub

Private Sub Form_Load()
SQ_OPER = 2
PUB_TIPREG = 600
PUB_CODCIA = "00"
LEER_TAB_LLAVE
Do Until tab_mayor.EOF
  If Val(tab_mayor!tab_codclie) <> 1 Then
      cmdtipo.AddItem Trim(tab_mayor!TAB_NOMLARGO) & String(50, " ") & (tab_mayor!TAB_NUMTAB)
  End If
  tab_mayor.MoveNext
Loop
'MSTDEF.DSN = PUB_DSN

pub_cadena = "SELECT * FROM COPARAM WHERE COP_CODCIA = ?"
Set PSCOP_LLAVE = CN.CreateQuery("", pub_cadena)
PSCOP_LLAVE(0) = 0
Set cop_llave = PSCOP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

PSCOP_LLAVE.rdoParameters(0) = LK_CODCIA
cop_llave.Requery
If cop_llave.EOF Then
 Screen.MousePointer = 0
 MsgBox "Hay que definir parametros de contabilidad.", 48, Pub_Titulo
'' Unload FORM_CONTA
 'Exit Sub
End If

CU = 0
BLOQUEA_TEXT
def_cp.AddItem "       "
def_cp.AddItem "C=Clientes"
def_cp.AddItem "P=Proveedores"

def_signo_ccm.AddItem "0"
def_signo_ccm.AddItem "1"
def_signo_ccm.AddItem "-1"

def_signo_arm.AddItem "0"
def_signo_arm.AddItem "1"
def_signo_arm.AddItem "-1"

def_signo_car.AddItem "0"
def_signo_car.AddItem "1"
def_signo_car.AddItem "-1"
def_signo_car.AddItem "2"

def_signo_caja.AddItem "0"
def_signo_caja.AddItem "1"
def_signo_caja.AddItem "-1"

 
PUB_CODCIA = "00"
LLENADOS FrmSub.def_tipdoc, 8
PUB_CODCIA = "00"
LLENADOS FrmSub.def_abreviado, 13
PUB_CODCIA = "00"
LLENADOS FrmSub.def_tipmov, 4
PUB_CODCIA = LK_CODCIA
LLENADOS FrmSub.def_calidad, 2

PUB_CODCIA = LK_CODCIA
LLENADOS FrmSub.def_codpro, 888

PUB_CODCIA = LK_CODCIA
LLENADOS FrmSub.def_art_gru, 122

LIMPIA_DEF
F1.Visible = True
cmdCancelar.Enabled = True

End Sub

Public Sub MENSAJE_DEF(TEXTO As String)
  LblMensaje.Caption = TEXTO
  PARPADEA.Enabled = True
End Sub

Public Sub ASIGNA(wcontrol As ComboBox, txt As String)
Dim C As Integer
For C = 0 To wcontrol.ListCount - 1
    If Trim(Left(wcontrol.List(C), 10)) = Trim(txt) Then
        wcontrol.ListIndex = C
        Exit Sub
    End If
Next C
End Sub
Public Sub ASIGNA_SIGNO(wcontrol As ComboBox, txt As Integer)
Dim C As Integer
For C = 0 To wcontrol.ListCount - 1
    If Val(wcontrol.List(C)) = txt Then
        wcontrol.ListIndex = C
        Exit Sub
    End If
Next C
End Sub
Public Sub ASIGNA_INT(wcontrol As ComboBox, txt As Integer)
Dim C As Integer
For C = 0 To wcontrol.ListCount - 1
    If Val(Trim(Right(wcontrol.List(C), 3))) = txt Then
        wcontrol.ListIndex = C
        Exit Sub
    End If
Next C
End Sub


Public Sub BLOQUEA_TEXT()
For i = 1 To 38
TxtDef(i).Enabled = False
Next i

def_cp.Enabled = False
def_signo_ccm.Enabled = False
def_signo_arm.Enabled = False
def_signo_car.Enabled = False
def_signo_caja.Enabled = False
def_tipmov.Enabled = False

def_is.Enabled = False
def_descto.Enabled = False
def_calidad.Enabled = False
def_tipdoc.Enabled = False
def_abreviado.Enabled = False
def_art_gru.Enabled = False

def_jabas.Enabled = False
def_codpro.Enabled = False
def_pollos.Enabled = False

def_mortal.Enabled = False
def_precio(1).Enabled = False
def_precio(2).Enabled = False
def_precio(3).Enabled = False
def_precio(4).Enabled = False

Frame2(1).Enabled = False
opt(0).Enabled = False
opt(1).Enabled = False
TxtDef(40).Enabled = False
TxtDef(41).Enabled = False
TxtDef(42).Enabled = False
txtsutgru.Enabled = False



End Sub
Public Sub DESBLOQUEA_TEXT()
For i = 1 To 38
    TxtDef(i).Enabled = True
Next i
If LK_CODUSU = "ADMIN" Then
 def_cp.Enabled = True
 def_signo_ccm.Enabled = True
 def_signo_car.Enabled = True
 def_signo_arm.Enabled = True
 def_signo_caja.Enabled = True
 def_tipmov.Enabled = True
 def_descto.Enabled = True
 def_calidad.Enabled = True
 def_is.Enabled = True
 def_tipdoc.Enabled = True
 def_abreviado.Enabled = True
 def_art_gru.Enabled = True
 def_jabas.Enabled = True
 def_codpro.Enabled = True
 def_pollos.Enabled = True
 def_mortal.Enabled = True
End If
Frame2(1).Enabled = True
def_precio(1).Enabled = True
def_precio(2).Enabled = True
def_precio(3).Enabled = True
def_precio(4).Enabled = True
opt(0).Enabled = True
opt(1).Enabled = True

TxtDef(40).Enabled = True
TxtDef(41).Enabled = True
TxtDef(42).Enabled = True
txtsutgru.Enabled = True

End Sub

Public Sub LIMPIA_DEF()
checrecont.Value = 0
cmdtipo.ListIndex = -1
Dim i As Integer
For i = 0 To 42
    TxtDef(i).text = ""
Next i
def_signo_ccm.ListIndex = -1
def_signo_car.ListIndex = -1
def_signo_arm.ListIndex = -1
def_signo_caja.ListIndex = -1

def_cp.ListIndex = -1
def_tipmov.ListIndex = -1
def_tipdoc.ListIndex = -1
def_abreviado.ListIndex = -1
def_calidad.ListIndex = -1
def_art_gru.ListIndex = -1
def_codpro.ListIndex = -1
def_descto.text = ""
def_is.text = ""
def_mortal.text = ""
'def_copia.Value = 0
def_precio(1).Value = False
def_precio(2).Value = False
def_precio(3).Value = False
def_precio(4).Value = False
Check1(0).Value = 0
Check1(1).Value = 0
Check1(2).Value = 0
Check1(3).Value = 0
Check1(4).Value = 0
txtsutgru.text = ""
def_jabas.text = ""
def_pollos.text = ""
def_concepto.text = ""
opt(0).Value = False
opt(1).Value = False

End Sub

Private Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    MSTDEF.Visible = False
    TxtDef(Index).SetFocus
End If
If KeyAscii = 13 And Index = 0 Then
    If Len(Grid1.TextMatrix(Grid1.Row, 1)) = 0 Then
       TxtDef(0).SetFocus
       Exit Sub
    End If
    PUB_CODTRA = Val(Left(Grid1.TextMatrix(Grid1.Row, 1), 4))
    MSTDEF.Visible = False
    LLENA_DEF 0
    cmdCancelar.Enabled = True
    TxtDef(IND).text = Grid1.TextMatrix(Grid1.Row, 1)
    TxtDef(IND + 2).text = Grid1.TextMatrix(Grid1.Row, 3)
    TxtDef(0).Locked = True
    cmdCopiar.Enabled = True
    chetransa.Enabled = True
    cmdconsultar.Enabled = True
    CmdModificar.SetFocus
End If


End Sub

Private Sub ListExiste_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    FrmSub.ListExiste.Visible = False
    FrmSub.TxtDef(1).SetFocus
    KeyAscii = 0
  End If

End Sub

Private Sub ListExiste_LostFocus()
 ListExiste.Visible = False
 TxtDef(1).SetFocus
 
End Sub

Private Sub LisTransa_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  LisTransa.Visible = False
  TxtDef(0).SetFocus
End If
If KeyAscii = 13 Then
  TxtDef(0).text = Trim(Left(LisTransa.text, 5))
  TxtDef_KeyPress 0, 13
  LisTransa.Visible = False
End If
End Sub

Private Sub opt_Click(Index As Integer)

  If Index = 0 Then
    FRAOPT.Left = 7800
    fracnt.Visible = True
  Else
    FRAOPT.Left = 3480
    fracnt.Visible = False
  End If
 
End Sub

Private Sub PARPADEA_Timer()
 CU = CU + 1
 LblMensaje.Visible = True 'Not LblMensaje.Visible
 If CU > 8 Then
   CU = 0
   PARPADEA.Enabled = False
   LblMensaje.Visible = False
 End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If PreviousTab = 0 Then
   FRAOPT.Visible = True
Else
  FRAOPT.Visible = False
End If
End Sub

Private Sub TxtDef_GotFocus(Index As Integer)
Select Case Index
   Case 0
      MSTDEF.Visible = False
   Case 5, 8, 11, 14, 17, 20, 23, 26, 29, 32, 35, 38
      Azul TxtDef(Index), TxtDef(Index)
End Select
End Sub
Private Sub TxtDef_KeyPress(Index As Integer, KeyAscii As Integer)
Dim car As String, tempo As Integer
car = Chr$(KeyAscii)
car = UCase$(Chr$(KeyAscii))
tempo = KeyAscii
KeyAscii = Asc(car)
If Index = 40 Then
  SOLO_ENTERO KeyAscii
End If
If Index = 3 Or Index = 6 Or Index = 9 Or Index = 12 Or Index = 15 Or Index = 18 Or Index = 21 Or Index = 24 Or Index = 27 Or Index = 30 Or Index = 33 Or Index = 36 Then
   If KeyAscii = 27 Then
     TxtDef(Index).text = ""
     TxtDef(Index).SetFocus
     Exit Sub
   End If
   If KeyAscii = 13 Then
     If Left(TxtDef(Index).text, 1) = "*" Then
       BUSCAR_CTA 0, Index
       Exit Sub
     End If
     Azul TxtDef(Index + 1), TxtDef(Index + 1)
   End If
   
End If
Select Case Index
Case 2
    KeyAscii = tempo
    If keyasci = 13 Then
     If TxtDef(Index + 1).Visible Then
       Azul TxtDef(Index + 1), TxtDef(Index + 1)
     Else
      opt(0).SetFocus
     End If
    End If
Case 0
   If Not car < "0" And car > "9" Then
      If KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
       End If
   End If
   If KeyAscii = 27 Then
      TxtDef(0).text = ""
      Exit Sub
   End If
   If KeyAscii = 13 Then
        If Trim(TxtDef(0).text) = "" Then
          Exit Sub
        End If
        If FrmSub.CmdModificar.Enabled = False Then
           Exit Sub
        End If
        If Val(TxtDef(0).text) = 0 Then Exit Sub
        IND = Index
        llave1 = ""
        UNICO = ""
        PUB_CODTRA = Val(TxtDef(0).text)
        KeyAscii = 0
        Call PROC_DEF
        If PUB_CODTRA <> -1 Then
         MSTDEF.Visible = True
         Grid1.Col = 3
         Grid1.Row = 1
         Grid1.SetFocus
        Else
         MsgBox "No existe Transacción...", 48, Pub_Titulo
         PUB_CODTRA = 0
         Azul TxtDef(0), TxtDef(0)
        End If
        Exit Sub
    End If
Case 1
   If Not car < "0" And car > "9" Then
      If KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
        
        Exit Sub
       End If
   End If

Case 4, 7, 10, 13, 16, 19, 22, 25, 28, 31, 34, 37
    If KeyAscii <> 68 And KeyAscii <> 72 And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 32 Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 13 Then
       Azul TxtDef(Index + 1), TxtDef(Index + 1)
       Exit Sub
    End If
Case 5, 8, 11, 14, 17, 20, 23, 26, 29, 32, 35, 38
  SOLO_ENTERO KeyAscii
   If KeyAscii = 13 Then
       Azul TxtDef(Index + 1), TxtDef(Index + 1)
       Exit Sub
   End If
End Select

End Sub

Public Sub LLENA_DEF(ban As Integer)
Dim i As Integer
If ban = 0 Then
    SQ_OPER = 1
    PUB_CODTRA = PUB_CODTRA
    PUB_SECUENCIA = Grid1.TextMatrix(Grid1.Row, 2)
    LEER_SUT_LLAVE
    SQ_OPER = 1
    PUB_CODTRA = Left(Grid1.TextMatrix(Grid1.Row, 1), 4)
    PUB_SECUENCIA = Grid1.TextMatrix(Grid1.Row, 2)
    PUB_CODCIA = LK_CODCIA
    LEER_CNT_LLAVE
End If
If Not cnt_llave.EOF Then
 For i = 3 To 38
     FrmSub.TxtDef(i).text = Trim(Nulo_Valors(cnt_llave.rdoColumns(i)))
 Next i
 TxtDef(40).text = Nulo_Valor0(cnt_llave!cnt_serie)
 TxtDef(41).text = Nulo_Valor0(cnt_llave!cnt_serie_recibo)
 TxtDef(42).text = Nulo_Valor0(cnt_llave!cnt_serie_caja)
End If
For i = 0 To 2
    FrmSub.TxtDef(i).text = Trim(Nulo_Valors(SUT_LLAVE.rdoColumns(i)))
Next i
If Trim(Nulo_Valors(SUT_LLAVE!SUT_cp)) = "C" Then
  def_cp.ListIndex = 1
ElseIf Trim(Nulo_Valors(SUT_LLAVE!SUT_cp)) = "P" Then
  def_cp.ListIndex = 2
Else
  def_cp.ListIndex = 0
End If
ASIGNA def_tipdoc, Trim(Nulo_Valors(SUT_LLAVE!SUT_tipdoc))
ASIGNA def_abreviado, Trim(Nulo_Valors(SUT_LLAVE!SUT_abreviado))

ASIGNA_SIGNO def_signo_ccm, Nulo_Valor0(SUT_LLAVE!SUT_signo_ccm)
ASIGNA_SIGNO def_signo_car, Nulo_Valor0(SUT_LLAVE!SUT_signo_car)
ASIGNA_SIGNO def_signo_arm, Nulo_Valor0(SUT_LLAVE!SUT_signo_arm)
ASIGNA_SIGNO def_signo_caja, Nulo_Valor0(SUT_LLAVE!SUT_signo_caja)
ASIGNA_INT def_tipmov, Nulo_Valor0(SUT_LLAVE!SUT_tipmov)
ASIGNA_INT cmdtipo, Nulo_Valor0(SUT_LLAVE!SUT_tipmov_REF)

ASIGNA_INT def_calidad, Nulo_Valor0(SUT_LLAVE!SUT_CALIDAD)
ASIGNA_INT def_codpro, Nulo_Valor0(SUT_LLAVE!SUT_codpro)
ASIGNA_INT def_art_gru, Nulo_Valor0(SUT_LLAVE!sut_art_gru)

def_mortal.text = Nulo_Valors(SUT_LLAVE!SUT_FLAG_TIPO)
'def_copia.Value = Val(Nulo_Valors(SUT_LLAVE!SUT_copia))
If Nulo_Valors(SUT_LLAVE!SUT_precio) = "1" Then 'A costo
 def_precio(1).Value = True
ElseIf Nulo_Valors(SUT_LLAVE!SUT_precio) = "2" Then  'Digitado
 def_precio(2).Value = True
ElseIf Nulo_Valors(SUT_LLAVE!SUT_precio) = "3" Then   'A Costo
 def_precio(3).Value = True
Else ' Fijo
 def_precio(4).Value = True
End If

Check1(0).Value = False
Check1(1).Value = False
Check1(2).Value = False
Check1(3).Value = False
Check1(4).Value = False

If Mid(SUT_LLAVE!SUT_PRECIOS, 1, 1) = "1" Then Check1(0).Value = 1
If Mid(SUT_LLAVE!SUT_PRECIOS, 2, 1) = "1" Then Check1(1).Value = 1
If Mid(SUT_LLAVE!SUT_PRECIOS, 3, 1) = "1" Then Check1(2).Value = 1
If Mid(SUT_LLAVE!SUT_PRECIOS, 4, 1) = "1" Then Check1(3).Value = 1
If Mid(SUT_LLAVE!SUT_PRECIOS, 5, 1) = "1" Then Check1(4).Value = 1

def_jabas.text = Trim(Nulo_Valors(SUT_LLAVE!SUT_jabas))
def_pollos.text = Trim(Nulo_Valors(SUT_LLAVE!SUT_DESFAC))
def_descto.text = Trim(Nulo_Valor0(SUT_LLAVE!sut_descto))
def_is.text = Trim(Nulo_Valors(SUT_LLAVE!sut_is))
txtsutgru.text = Trim(Nulo_Valors(SUT_LLAVE!sut_art_gru))
checrecont.Value = Trim(Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC))

opt(0).Value = False
opt(1).Value = False
If Trim(Nulo_Valors(SUT_LLAVE!SUT_FLAG_CONTABLE)) = "A" Then
  opt(0).Value = True
  FRAOPT.Left = 7800
  fracnt.Visible = True
Else
  opt(1).Value = True
  FRAOPT.Left = 3480
  fracnt.Visible = False
End If

End Sub

Public Function EXISTE_DEF(VALOR1 As Integer, VALOR2 As String) As Boolean
ListExiste.Clear
FrmSub.ListExiste.Cols = 3
FrmSub.ListExiste.Row = 0
FrmSub.ListExiste.Col = 0
FrmSub.ListExiste.text = " TRANS "
FrmSub.ListExiste.Col = 1
FrmSub.ListExiste.text = " SECUENCIA "
FrmSub.ListExiste.Col = 2
FrmSub.ListExiste.text = " DESCRIPCION "
FrmSub.ListExiste.ColWidth(0) = 1000
FrmSub.ListExiste.ColWidth(1) = 1200
FrmSub.ListExiste.ColWidth(2) = 6000
FrmSub.ListExiste.Left = 120
FrmSub.ListExiste.Top = 3720
FrmSub.ListExiste.Height = 1335
FrmSub.ListExiste.Width = 9255
EXISTE_DEF = False

Dim NUMCAMPO As Integer
Dim OJO As String * 1
Static P As Boolean
Dim VAR As String
Dim chec1 As Boolean
Dim chec2 As Boolean
Dim chec3 As Boolean
If Not VALOR2 = "" Then
SQ_OPER = 1
PUB_CODTRA = VALOR1
PUB_SECUENCIA = Val(VALOR2)
LEER_SUT_LLAVE
If SUT_LLAVE.EOF Then
'   msgbox "NO HAY DATOS... "
   GoTo FIN
End If
End If
SQ_OPER = 2
PUB_CODTRA = VALOR1
LEER_SUT_LLAVE
If SUT_MAYOR.EOF Then
'   msgbox "NO HAY DATOS... "
   GoTo FIN
End If
SUT_MAYOR.MoveFirst
fila = 0
FrmSub.ListExiste.Rows = 2
Do Until SUT_MAYOR.EOF 'Or fila = 50
    fila = fila + 1
    FrmSub.ListExiste.TextMatrix(fila, 0) = Nulo_Valors(SUT_MAYOR!DEF_CODTRA)
    FrmSub.ListExiste.TextMatrix(fila, 1) = Nulo_Valors(SUT_MAYOR!DEF_SECUENCIA)
    FrmSub.ListExiste.TextMatrix(fila, 2) = Nulo_Valors(SUT_MAYOR!DEF_DESCRIPCION)
    EXISTE_DEF = True
    FrmSub.ListExiste.Rows = FrmSub.ListExiste.Rows + 1
    SUT_MAYOR.MoveNext
Loop
If EXISTE_DEF Then
    FrmSub.ListExiste.TextMatrix(fila + 1, 0) = "*********"
    FrmSub.ListExiste.TextMatrix(fila + 1, 2) = "<ESC> Para salir. . ."
    FrmSub.ListExiste.Visible = True
    FrmSub.ListExiste.Row = 1
    FrmSub.ListExiste.Col = 1
    FrmSub.ListExiste.SetFocus
    
End If

GoTo FIN
Exit Function

CHECKERROR:
MsgBox Err.Description
FIN:

End Function

Public Function CONSIS_CUENTA() As Integer
cop_llave.Requery
If cop_llave.EOF Then
  CONSIS_CUENTA = 0
  Exit Function
End If
SQ_OPER = 1
For fila = 3 To 36 Step 3
  If Trim(TxtDef(fila).text) <> "" Then
     If Trim(TxtDef(fila).text) = "FACART" Or Trim(TxtDef(fila).text) = "CLIENTES" Or Trim(TxtDef(fila).text) = "CLIENTES2" Then
          GoTo VALE
     End If
     PUB_CUENTA = Trim(TxtDef(fila).text)
     LEER_COM_LLAVE
     If com_llave.EOF Then
        MsgBox "Cuenta,  NO EXISTE ... ", 48, Pub_Titulo
        Azul TxtDef(fila), TxtDef(fila)
        CONSIS_CUENTA = 1
        Exit Function
     ElseIf Val(com_llave!COM_NIVEL) <> Nulo_Valor0(cop_llave!COP_NIVEL_AFECTACION) Then
        MsgBox "Cuenta,  Invalidad ... no procede", 48, Pub_Titulo
        Azul TxtDef(fila), TxtDef(fila)
        CONSIS_CUENTA = 1
        Exit Function
     End If
VALE:
  End If
Next fila
CONSIS_CUENTA = 0
End Function

Private Sub TxtDef_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If Index = 0 Then
 If KeyCode = 112 Then
   LLENA_LISTRANSA FrmSub.LisTransa
   FrmSub.LisTransa.Left = 980
   FrmSub.LisTransa.Top = 0
   FrmSub.LisTransa.Width = 4500
   FrmSub.LisTransa.Height = 3000
   FrmSub.LisTransa.SetFocus
 End If

End If

End Sub
Public Sub BUSCAR_CTA(WTIPO As Integer, windex As Integer)
Dim wcuenta As TextBox
Dim wgrupo As String
Dim wq_cuenta As String

LK_TABLA = "SUB_TRA"
If WTIPO = 1 Then
 If TxtDef(windex).text = "*" Then
  wgrupo = "" 'Trim(i_cuenta.text)
  archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_CUENTA >= '" & wgrupo & "'  ORDER BY COM_CUENTA"
 Else
 TxtDef(windex).text = Mid(TxtDef(windex).text, 2, Len(TxtDef(windex).text))
 wgrupo = Trim(TxtDef(windex).text)
 If Val(wgrupo) = 0 Then Exit Sub
 archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_CUENTA >= '" & wgrupo & "' AND COM_CUENTA < '" & Trim(str(Val(wgrupo) + 1)) & "'  ORDER BY COM_CUENTA"
 End If
Else
 If TxtDef(windex).text = "*" Then
  wgrupo = "" 'Trim(i_cuenta.text)
  archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_CUENTA >= '" & wgrupo & "'  ORDER BY COM_CUENTA"
 Else
 TxtDef(windex).text = Mid(TxtDef(windex).text, 2, Len(TxtDef(windex).text))
 wgrupo = Trim(TxtDef(windex).text)
 If Val(wgrupo) = 0 Then Exit Sub
 archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_CUENTA >= '" & wgrupo & "' AND COM_CUENTA < '" & Trim(str(Val(wgrupo) + 1)) & "'  ORDER BY COM_CUENTA"
 End If
End If
Load frmBuscacta
frmBuscacta.lbltabla.Caption = LK_TABLA
frmBuscacta.Show 1
wq_cuenta = Trim(frmBuscacta.tcuenta)
If wq_cuenta <> "" Then
  If WTIPO = 1 Then
    i_cuenta.text = Trim(frmBuscacta.tcuenta)
    lcuenta.Caption = Trim(frmBuscacta.tnombre.text)
  Else
  'textovar.text = Trim(frmBuscacta.tcuenta)
  TxtDef(windex).text = Trim(frmBuscacta.tcuenta)
  End If
End If
Unload frmBuscacta
If wq_cuenta <> "" Then
   If WTIPO = 1 Then
      TxtDef_KeyPress windex, 13
      'i_cuenta_KeyPress 13
   Else
     TxtDef_KeyPress windex, 13
   End If
ElseIf wq_cuenta <> "" Then
  If WTIPO = 1 Then
     TxtDef_KeyPress windex, 13
  Else
      TxtDef_KeyPress windex, 13
  End If
Else
  ''''''Azul3 textovar, textovar
End If


End Sub


Public Function CONSIS() As Boolean
Dim wfa

CONSIS = True
wfa = ""
PUB_CUENTA = ""
If Trim(TxtDef(3)) <> "" Then
    PUB_CUENTA = Trim(TxtDef(3))
    pu_codcia = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
      If Trim(PUB_CUENTA) = "BANCOS" Or Trim(PUB_CUENTA) = "BANCOS2" Or Trim(PUB_CUENTA) = "CLIENTES" Or Right(Trim(PUB_CUENTA), 2) = "XX" Or Trim(PUB_CUENTA) = "CLIENTES2" Then
        GoTo salta1
      Else
        wfa = "X"
        Azul TxtDef(3), TxtDef(3)
        GoTo SALIG
      End If
    End If
    If Val(com_llave!COM_NIVEL) <> Val(cop_llave!cop_nivel_max) Then
      wfa = "Y"
      Azul TxtDef(3), TxtDef(3)
      GoTo SALIG
    End If
salta1:
    If Trim(TxtDef(4)) = "" Then
       wfa = "D"
       Azul TxtDef(4), TxtDef(4)
       GoTo SALIG
    End If
    If Val(TxtDef(5)) = 0 Then
       wfa = "F"
       Azul TxtDef(5), TxtDef(5)
       GoTo SALIG
    End If
End If
If Trim(TxtDef(6)) <> "" Then
    PUB_CUENTA = Trim(TxtDef(6))
    pu_codcia = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
      If Trim(PUB_CUENTA) = "BANCOS" Or Trim(PUB_CUENTA) = "BANCOS2" Or Trim(PUB_CUENTA) = "CLIENTES" Or Right(Trim(PUB_CUENTA), 2) = "XX" Or Trim(PUB_CUENTA) = "CLIENTES2" Then
      GoTo salta2
      Else
        wfa = "X"
        Azul TxtDef(6), TxtDef(6)
        GoTo SALIG
      End If
    End If
    If Val(com_llave!COM_NIVEL) <> Val(cop_llave!cop_nivel_max) Then
      wfa = "Y"
      Azul TxtDef(6), TxtDef(6)
      GoTo SALIG
    End If
salta2:
    If Trim(TxtDef(7)) = "" Then
       wfa = "D"
       Azul TxtDef(7), TxtDef(7)
       GoTo SALIG
    End If
    If Val(TxtDef(8)) = 0 Then
       wfa = "F"
       Azul TxtDef(8), TxtDef(8)
       GoTo SALIG
    End If
End If
If Trim(TxtDef(9)) <> "" Then
    PUB_CUENTA = Trim(TxtDef(9))
    pu_codcia = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
      If Trim(PUB_CUENTA) = "BANCOS" Or Trim(PUB_CUENTA) = "BANCOS2" Or Trim(PUB_CUENTA) = "CLIENTES" Or Right(Trim(PUB_CUENTA), 2) = "XX" Or Trim(PUB_CUENTA) = "CLIENTES2" Then
         GoTo salta3
      Else
        wfa = "X"
        Azul TxtDef(9), TxtDef(9)
        GoTo SALIG
      End If
    End If
    If Val(com_llave!COM_NIVEL) <> Val(cop_llave!cop_nivel_max) Then
      wfa = "Y"
      Azul TxtDef(9), TxtDef(9)
      GoTo SALIG
    End If
salta3:
    If Trim(TxtDef(10)) = "" Then
       wfa = "D"
       Azul TxtDef(10), TxtDef(10)
       GoTo SALIG
    End If
    If Val(TxtDef(11)) = 0 Then
       wfa = "F"
       Azul TxtDef(11), TxtDef(11)
       GoTo SALIG
    End If
End If
If Trim(TxtDef(12)) <> "" Then
    PUB_CUENTA = Trim(TxtDef(12))
    pu_codcia = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
      If Trim(PUB_CUENTA) = "BANCOS" Or Trim(PUB_CUENTA) = "BANCOS2" Or Trim(PUB_CUENTA) = "CLIENTES" Or Right(Trim(PUB_CUENTA), 2) = "XX" Or Trim(PUB_CUENTA) = "CLIENTES2" Then
         GoTo salta4
      Else
        wfa = "X"
        Azul TxtDef(12), TxtDef(12)
        GoTo SALIG
      End If
    End If
    If Val(com_llave!COM_NIVEL) <> Val(cop_llave!cop_nivel_max) Then
      If Trim(PUB_CUENTA) = "BANCOS" Or Trim(PUB_CUENTA) = "BANCOS2" Or Trim(PUB_CUENTA) = "CLIENTES" Or Right(Trim(PUB_CUENTA), 2) = "XX" Then
      Else
        wfa = "Y"
        Azul TxtDef(12), TxtDef(12)
        GoTo SALIG
      End If
    End If
salta4:
    If Trim(TxtDef(13)) = "" Then
       wfa = "D"
       Azul TxtDef(13), TxtDef(13)
       GoTo SALIG
    End If
    If Val(TxtDef(14)) = 0 Then
       wfa = "F"
       Azul TxtDef(14), TxtDef(14)
       GoTo SALIG
    End If
End If
If Trim(TxtDef(15)) <> "" Then
    PUB_CUENTA = Trim(TxtDef(15))
    pu_codcia = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
      If Trim(PUB_CUENTA) = "BANCOS" Or Trim(PUB_CUENTA) = "BANCOS2" Or Trim(PUB_CUENTA) = "CLIENTES" Or Right(Trim(PUB_CUENTA), 2) = "XX" Or Trim(PUB_CUENTA) = "CLIENTES2" Then
          GoTo salta5
      Else
        wfa = "X"
        Azul TxtDef(15), TxtDef(15)
        GoTo SALIG
      End If
    End If
    If Val(com_llave!COM_NIVEL) <> Val(cop_llave!cop_nivel_max) Then
      wfa = "Y"
      Azul TxtDef(15), TxtDef(15)
      GoTo SALIG
    End If
salta5:
    If Trim(TxtDef(16)) = "" Then
       wfa = "D"
       Azul TxtDef(16), TxtDef(16)
       GoTo SALIG
    End If
    If Val(TxtDef(17)) = 0 Then
       wfa = "F"
       Azul TxtDef(17), TxtDef(17)
       GoTo SALIG
    End If
End If
If Trim(TxtDef(18)) <> "" Then
    PUB_CUENTA = Trim(TxtDef(18))
    pu_codcia = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
      If Trim(PUB_CUENTA) = "BANCOS" Or Trim(PUB_CUENTA) = "BANCOS2" Or Trim(PUB_CUENTA) = "CLIENTES" Or Right(Trim(PUB_CUENTA), 2) = "XX" Or Trim(PUB_CUENTA) = "CLIENTES2" Then
         GoTo salta6
      Else
        wfa = "X"
        Azul TxtDef(18), TxtDef(18)
        GoTo SALIG
      End If
    End If
    If Val(com_llave!COM_NIVEL) <> Val(cop_llave!cop_nivel_max) Then
      wfa = "Y"
      Azul TxtDef(18), TxtDef(18)
      GoTo SALIG
    End If
salta6:
    If Trim(TxtDef(19)) = "" Then
       wfa = "D"
       Azul TxtDef(19), TxtDef(19)
       GoTo SALIG
    End If
    If Val(TxtDef(20)) = 0 Then
       wfa = "F"
       Azul TxtDef(20), TxtDef(20)
       GoTo SALIG
    End If
End If
If Trim(TxtDef(21)) <> "" Then
    PUB_CUENTA = Trim(TxtDef(21))
    pu_codcia = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
      If Trim(PUB_CUENTA) = "BANCOS" Or Trim(PUB_CUENTA) = "BANCOS2" Or Trim(PUB_CUENTA) = "CLIENTES" Or Right(Trim(PUB_CUENTA), 2) = "XX" Or Trim(PUB_CUENTA) = "CLIENTES2" Then
         GoTo salta7
      Else
        wfa = "X"
        Azul TxtDef(21), TxtDef(21)
        GoTo SALIG
      End If
    End If
    If Val(com_llave!COM_NIVEL) <> Val(cop_llave!cop_nivel_max) Then
      wfa = "Y"
      Azul TxtDef(21), TxtDef(21)
      GoTo SALIG
    End If
salta7:
    If Trim(TxtDef(22)) = "" Then
       wfa = "D"
       Azul TxtDef(22), TxtDef(22)
       GoTo SALIG
    End If
    If Val(TxtDef(23)) = 0 Then
       wfa = "F"
       Azul TxtDef(23), TxtDef(23)
       GoTo SALIG
    End If

End If
If Trim(TxtDef(24)) <> "" Then
    PUB_CUENTA = Trim(TxtDef(24))
    pu_codcia = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
      If Trim(PUB_CUENTA) = "BANCOS" Or Trim(PUB_CUENTA) = "BANCOS2" Or Trim(PUB_CUENTA) = "CLIENTES" Or Right(Trim(PUB_CUENTA), 2) = "XX" Or Trim(PUB_CUENTA) = "CLIENTES2" Then
      GoTo salta8
      Else
        wfa = "X"
        Azul TxtDef(24), TxtDef(24)
        GoTo SALIG
      End If
    End If
    If Val(com_llave!COM_NIVEL) <> Val(cop_llave!cop_nivel_max) Then
      wfa = "Y"
      Azul TxtDef(24), TxtDef(24)
      GoTo SALIG
    End If
salta8:
    If Trim(TxtDef(25)) = "" Then
       wfa = "D"
       Azul TxtDef(25), TxtDef(25)
       GoTo SALIG
    End If
    If Val(TxtDef(26)) = 0 Then
       wfa = "F"
       Azul TxtDef(26), TxtDef(26)
       GoTo SALIG
    End If

End If
If Trim(TxtDef(27)) <> "" Then
    PUB_CUENTA = Trim(TxtDef(27))
    pu_codcia = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
      If Trim(PUB_CUENTA) = "BANCOS" Or Trim(PUB_CUENTA) = "BANCOS2" Or Trim(PUB_CUENTA) = "CLIENTES" Or Right(Trim(PUB_CUENTA), 2) = "XX" Or Trim(PUB_CUENTA) = "CLIENTES2" Then
      GoTo salta9
      Else
        wfa = "X"
        Azul TxtDef(27), TxtDef(27)
        GoTo SALIG
      End If
    End If
    If Val(com_llave!COM_NIVEL) <> Val(cop_llave!cop_nivel_max) Then
      wfa = "Y"
      Azul TxtDef(27), TxtDef(27)
      GoTo SALIG
    End If
salta9:
    If Trim(TxtDef(28)) = "" Then
       wfa = "D"
       Azul TxtDef(28), TxtDef(28)
       GoTo SALIG
    End If
    If Val(TxtDef(29)) = 0 Then
       wfa = "F"
       Azul TxtDef(29), TxtDef(29)
       GoTo SALIG
    End If

End If
If Trim(TxtDef(30)) <> "" Then
    PUB_CUENTA = Trim(TxtDef(30))
    pu_codcia = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
      If Trim(PUB_CUENTA) = "BANCOS" Or Trim(PUB_CUENTA) = "BANCOS2" Or Trim(PUB_CUENTA) = "CLIENTES" Or Right(Trim(PUB_CUENTA), 2) = "XX" Or Trim(PUB_CUENTA) = "CLIENTES2" Then
         GoTo salta10
      Else
        wfa = "X"
        Azul TxtDef(30), TxtDef(30)
        GoTo SALIG
      End If
    End If
    If Val(com_llave!COM_NIVEL) <> Val(cop_llave!cop_nivel_max) Then
      wfa = "Y"
      Azul TxtDef(30), TxtDef(30)
      GoTo SALIG
    End If
salta10:
    If Trim(TxtDef(31)) = "" Then
       wfa = "D"
       Azul TxtDef(31), TxtDef(31)
       GoTo SALIG
    End If
    If Val(TxtDef(32)) = 0 Then
       wfa = "F"
       Azul TxtDef(32), TxtDef(32)
       GoTo SALIG
    End If

End If
If Trim(TxtDef(33)) <> "" Then
    PUB_CUENTA = Trim(TxtDef(33))
    pu_codcia = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
      If Trim(PUB_CUENTA) = "BANCOS" Or Trim(PUB_CUENTA) = "BANCOS2" Or Trim(PUB_CUENTA) = "CLIENTES" Or Right(Trim(PUB_CUENTA), 2) = "XX" Or Trim(PUB_CUENTA) = "CLIENTES2" Then
      GoTo salta11
      Else
        wfa = "X"
        Azul TxtDef(33), TxtDef(33)
        GoTo SALIG
      End If
    End If
    If Val(com_llave!COM_NIVEL) <> Val(cop_llave!cop_nivel_max) Then
      wfa = "Y"
      Azul TxtDef(33), TxtDef(33)
      GoTo SALIG
    End If
salta11:
    If Trim(TxtDef(34)) = "" Then
       wfa = "D"
       Azul TxtDef(34), TxtDef(34)
       GoTo SALIG
    End If
    If Val(TxtDef(35)) = 0 Then
       wfa = "F"
       Azul TxtDef(35), TxtDef(35)
       GoTo SALIG
    End If

End If
If Trim(TxtDef(36)) <> "" Then
    PUB_CUENTA = Trim(TxtDef(36))
    pu_codcia = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
      If Trim(PUB_CUENTA) = "BANCOS" Or Trim(PUB_CUENTA) = "BANCOS2" Or Trim(PUB_CUENTA) = "CLIENTES" Or Right(Trim(PUB_CUENTA), 2) = "XX" Or Trim(PUB_CUENTA) = "CLIENTES2" Then
      GoTo salta13
      Else
        wfa = "X"
        Azul TxtDef(36), TxtDef(36)
        GoTo SALIG
      End If
    End If
    If Val(com_llave!COM_NIVEL) <> Val(cop_llave!cop_nivel_max) Then
      wfa = "Y"
      Azul TxtDef(36), TxtDef(36)
      GoTo SALIG
    End If
salta13:
    If Trim(TxtDef(37)) = "" Then
       wfa = "D"
       Azul TxtDef(37), TxtDef(37)
       GoTo SALIG
    End If
    If Val(TxtDef(38)) = 0 Then
       wfa = "F"
       Azul TxtDef(38), TxtDef(38)
       GoTo SALIG
    End If

End If

If opt(0).Value = True And Trim(PUB_CUENTA) = "" Then
   pub_mensaje = "No ha Definido sus Asientos Contables...   ¿Desea Continuar... ?"
   Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
   If Pub_Respuesta = vbYes Then
       CONSIS = True
   Else
       CONSIS = False
   End If
End If


Exit Function
SALIG:
If opt(0).Value = True Then
 If wfa = "X" Then
  MsgBox "Cuenta No existe en Plan de Cuenta ", 48, Pub_Titulo
 ElseIf wfa = "Y" Then
  MsgBox "Cuenta No es ultimo Nivel ", 48, Pub_Titulo
 ElseIf wfa = "D" Then
  MsgBox "No tiene Destino, Para el Debe = D , Para el Haber = H ", 48, Pub_Titulo
 ElseIf wfa = "F" Then
  MsgBox "No tiene Codigo numerico,Normalmente solo para Bancos es codigo = 8 , el resto de 7 ", 48, Pub_Titulo
 End If
CONSIS = False
End If
End Function
