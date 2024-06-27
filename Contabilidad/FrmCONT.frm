VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCont 
   Caption         =   "FrmCont"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9225
   Icon            =   "FrmCONT.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   9225
   WindowState     =   2  'Maximized
   Begin VB.Frame FCUENTAS 
      Caption         =   "Cuentas Contables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3375
      Left            =   4080
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   3735
      Begin VB.ListBox CUENTAS 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   3255
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Consultar 
      Height          =   735
      Left            =   5520
      TabIndex        =   16
      Top             =   5880
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1296
      _Version        =   327680
      Cols            =   4
      FixedCols       =   3
   End
   Begin MSFlexGridLib.MSFlexGrid ListExiste 
      Height          =   495
      Left            =   840
      TabIndex        =   11
      Top             =   6120
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      _Version        =   327680
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   400
      Left            =   5640
      TabIndex        =   3
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
      Left            =   4080
      TabIndex        =   9
      Text            =   " "
      Top             =   6480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame F1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   9375
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
         Left            =   240
         MaxLength       =   4
         TabIndex        =   138
         Text            =   "Def_codtra"
         Top             =   480
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
         Index           =   1
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   137
         Text            =   "def_secuencia"
         Top             =   480
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
         Index           =   2
         Left            =   3360
         MaxLength       =   30
         TabIndex        =   100
         Text            =   "def_descripcion"
         Top             =   480
         Width           =   4455
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   120
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Ce&rrar"
      Height          =   400
      Left            =   7440
      TabIndex        =   4
      Top             =   5280
      Width           =   1500
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   400
      Left            =   240
      TabIndex        =   0
      Top             =   5280
      Width           =   1500
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   400
      Left            =   3840
      TabIndex        =   2
      Top             =   5280
      Width           =   1500
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Height          =   400
      Left            =   2040
      TabIndex        =   1
      Top             =   5280
      Width           =   1500
   End
   Begin VB.Frame MSTDEF 
      Caption         =   "Seleccione  :"
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
      Left            =   5400
      TabIndex        =   14
      Top             =   4560
      Visible         =   0   'False
      Width           =   3975
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   3375
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   5953
         _Version        =   327680
         Rows            =   50
         Cols            =   4
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Index           =   0
      Left            =   0
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1080
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7011
      _Version        =   327680
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Datos (1)"
      TabPicture(0)   =   "FrmCONT.frx":0442
      Tab(0).ControlCount=   2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2(1)"
      Tab(0).Control(1).Enabled=   0   'False
      TabCaption(1)   =   "Datos (2)"
      TabPicture(1)   =   "FrmCONT.frx":045E
      Tab(1).ControlCount=   19
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(26)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(25)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(24)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(20)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(19)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(18)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(17)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(16)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label1(15)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(14)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label1(13)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label1(11)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label1(10)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label1(9)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label1(12)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label1(21)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label1(22)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label1(23)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "SSTab1(5)"
      Tab(1).Control(18).Enabled=   0   'False
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   -74880
         TabIndex        =   24
         Top             =   3120
         Width           =   4335
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
            Left            =   1440
            TabIndex        =   27
            Top             =   120
            Width           =   1455
         End
         Begin VB.CheckBox chetransa 
            Caption         =   "Incluir Transacción"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   160
            Width           =   1215
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
            Left            =   3120
            TabIndex        =   25
            Top             =   210
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "P r e c i o s"
         Height          =   1575
         Index           =   1
         Left            =   -67440
         TabIndex        =   18
         Top             =   2040
         Width           =   1335
         Begin VB.CheckBox Check1 
            Caption         =   "Precio 1"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Precio 2"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Precio 3"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Precio 4"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   20
            Top             =   960
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Precio 5"
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   19
            Top             =   1200
            Width           =   1095
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3975
         Index           =   5
         Left            =   0
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   0
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   7011
         _Version        =   327680
         Tabs            =   1
         TabHeight       =   520
         TabCaption(0)   =   "Datos (1)"
         TabPicture(0)   =   "FrmCONT.frx":047A
         Tab(0).ControlCount=   54
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1(32)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label1(31)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label1(30)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label1(29)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label1(28)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label1(6)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label1(38)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label1(37)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label1(36)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label1(35)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label1(34)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label1(33)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label1(27)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label1(8)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Label1(7)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Label1(5)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Label1(4)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Label1(3)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "TxtDef(3)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "TxtDef(4)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "TxtDef(5)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "TxtDef(6)"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "TxtDef(7)"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "TxtDef(8)"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "TxtDef(9)"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "TxtDef(10)"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "TxtDef(11)"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "TxtDef(12)"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "TxtDef(13)"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "TxtDef(14)"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "TxtDef(15)"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "TxtDef(16)"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "TxtDef(17)"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "TxtDef(18)"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "TxtDef(19)"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "TxtDef(20)"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "TxtDef(21)"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "TxtDef(22)"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).Control(38)=   "TxtDef(23)"
         Tab(0).Control(38).Enabled=   0   'False
         Tab(0).Control(39)=   "TxtDef(24)"
         Tab(0).Control(39).Enabled=   0   'False
         Tab(0).Control(40)=   "TxtDef(25)"
         Tab(0).Control(40).Enabled=   0   'False
         Tab(0).Control(41)=   "TxtDef(26)"
         Tab(0).Control(41).Enabled=   0   'False
         Tab(0).Control(42)=   "TxtDef(27)"
         Tab(0).Control(42).Enabled=   0   'False
         Tab(0).Control(43)=   "TxtDef(28)"
         Tab(0).Control(43).Enabled=   0   'False
         Tab(0).Control(44)=   "TxtDef(29)"
         Tab(0).Control(44).Enabled=   0   'False
         Tab(0).Control(45)=   "TxtDef(30)"
         Tab(0).Control(45).Enabled=   0   'False
         Tab(0).Control(46)=   "TxtDef(31)"
         Tab(0).Control(46).Enabled=   0   'False
         Tab(0).Control(47)=   "TxtDef(32)"
         Tab(0).Control(47).Enabled=   0   'False
         Tab(0).Control(48)=   "TxtDef(33)"
         Tab(0).Control(48).Enabled=   0   'False
         Tab(0).Control(49)=   "TxtDef(34)"
         Tab(0).Control(49).Enabled=   0   'False
         Tab(0).Control(50)=   "TxtDef(35)"
         Tab(0).Control(50).Enabled=   0   'False
         Tab(0).Control(51)=   "TxtDef(36)"
         Tab(0).Control(51).Enabled=   0   'False
         Tab(0).Control(52)=   "TxtDef(37)"
         Tab(0).Control(52).Enabled=   0   'False
         Tab(0).Control(53)=   "TxtDef(38)"
         Tab(0).Control(53).Enabled=   0   'False
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
            Left            =   6240
            MaxLength       =   8
            TabIndex        =   136
            Text            =   "def_campo12"
            Top             =   3240
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
            Left            =   5640
            MaxLength       =   1
            TabIndex        =   135
            Text            =   "def_dh12"
            Top             =   3240
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
            Left            =   4440
            MaxLength       =   12
            TabIndex        =   134
            Text            =   "def_cta12"
            Top             =   3240
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
            Left            =   6240
            MaxLength       =   8
            TabIndex        =   133
            Text            =   "def_campo11"
            Top             =   2760
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
            Left            =   5640
            MaxLength       =   1
            TabIndex        =   132
            Text            =   "def_dh11"
            Top             =   2760
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
            Left            =   4440
            MaxLength       =   12
            TabIndex        =   131
            Text            =   "def_cta11"
            Top             =   2760
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
            Left            =   6240
            MaxLength       =   8
            TabIndex        =   130
            Text            =   "def_campo10"
            Top             =   2280
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
            Left            =   5640
            MaxLength       =   1
            TabIndex        =   129
            Text            =   "def_dh10"
            Top             =   2280
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
            Left            =   4440
            MaxLength       =   12
            TabIndex        =   128
            Text            =   "def_cta10"
            Top             =   2280
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
            Left            =   6240
            MaxLength       =   8
            TabIndex        =   127
            Text            =   "def_campo9"
            Top             =   1800
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
            Left            =   5640
            MaxLength       =   1
            TabIndex        =   126
            Text            =   "def_dh9"
            Top             =   1800
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
            Left            =   4440
            MaxLength       =   12
            TabIndex        =   125
            Text            =   "def_cta9"
            Top             =   1800
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
            Left            =   6240
            MaxLength       =   8
            TabIndex        =   124
            Text            =   "def_campo8"
            Top             =   1320
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
            Left            =   5640
            MaxLength       =   1
            TabIndex        =   123
            Text            =   "def_dh8"
            Top             =   1320
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
            Left            =   4440
            MaxLength       =   12
            TabIndex        =   122
            Text            =   "def_cta8"
            Top             =   1320
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
            Left            =   6240
            MaxLength       =   8
            TabIndex        =   121
            Text            =   "def_campo7"
            Top             =   840
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
            Left            =   5640
            MaxLength       =   1
            TabIndex        =   120
            Text            =   "def_dh7"
            Top             =   840
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
            Left            =   4440
            MaxLength       =   12
            TabIndex        =   119
            Text            =   "def_cta7"
            Top             =   840
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
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   118
            Text            =   "def_campo6"
            Top             =   3240
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
            Left            =   1800
            MaxLength       =   1
            TabIndex        =   117
            Text            =   "def_dh6"
            Top             =   3240
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
            Left            =   600
            MaxLength       =   12
            TabIndex        =   116
            Text            =   "def_cta6"
            Top             =   3240
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
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   115
            Text            =   "def_campo5"
            Top             =   2760
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
            Left            =   1800
            MaxLength       =   1
            TabIndex        =   114
            Text            =   "def_dh5"
            Top             =   2760
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
            Left            =   600
            MaxLength       =   12
            TabIndex        =   113
            Text            =   "def_cta5"
            Top             =   2760
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
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   112
            Text            =   "def_campo4"
            Top             =   2280
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
            Left            =   1800
            MaxLength       =   1
            TabIndex        =   111
            Text            =   "def_dh4"
            Top             =   2280
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
            Left            =   600
            MaxLength       =   12
            TabIndex        =   110
            Text            =   "def_cta4"
            Top             =   2280
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
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   109
            Text            =   "def_campo3"
            Top             =   1800
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
            Left            =   1800
            MaxLength       =   1
            TabIndex        =   108
            Text            =   "def_dh3"
            Top             =   1800
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
            Left            =   600
            MaxLength       =   12
            TabIndex        =   107
            Text            =   "def_cta3"
            Top             =   1800
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
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   106
            Text            =   "def_campo2"
            Top             =   1320
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
            Left            =   1800
            MaxLength       =   1
            TabIndex        =   105
            Text            =   "def_dh2"
            Top             =   1320
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
            Left            =   600
            MaxLength       =   12
            TabIndex        =   104
            Text            =   "def_cta2"
            Top             =   1320
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
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   103
            Text            =   "def_campo1"
            Top             =   840
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
            Left            =   1800
            MaxLength       =   1
            TabIndex        =   102
            Text            =   "def_dh1"
            Top             =   840
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
            Left            =   600
            MaxLength       =   12
            TabIndex        =   101
            Text            =   "def_cta1"
            Top             =   840
            Width           =   975
         End
         Begin VB.Frame Frame3 
            Height          =   615
            Left            =   -74880
            TabIndex        =   61
            Top             =   3120
            Width           =   4335
            Begin VB.CommandButton Command1 
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
               Left            =   1440
               TabIndex        =   64
               Top             =   210
               Width           =   1455
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Incluir Transacción"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   63
               Top             =   160
               Width           =   1215
            End
            Begin VB.CommandButton Command2 
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
               Left            =   3120
               TabIndex        =   62
               Top             =   210
               Width           =   1095
            End
         End
         Begin VB.CheckBox def_mortal 
            Caption         =   "Concepto E/S"
            Height          =   255
            Left            =   -68040
            TabIndex        =   60
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CheckBox def_copia 
            Caption         =   "Digitación Totales"
            Height          =   375
            Left            =   -68040
            TabIndex        =   59
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Frame Frame2 
            Caption         =   "P r e c i o s"
            Height          =   975
            Index           =   0
            Left            =   -70320
            TabIndex        =   55
            Top             =   2760
            Width           =   2415
            Begin VB.OptionButton def_precio 
               Caption         =   "Por Articulo"
               Height          =   255
               Index           =   3
               Left            =   1080
               TabIndex        =   58
               Top             =   120
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton def_precio 
               Caption         =   "Digitado"
               Height          =   255
               Index           =   2
               Left            =   1080
               TabIndex        =   57
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton def_precio 
               Caption         =   "A Costo"
               Height          =   375
               Index           =   1
               Left            =   1080
               TabIndex        =   56
               Top             =   600
               Width           =   1095
            End
         End
         Begin VB.ComboBox def_tipmov_ref 
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
            Left            =   -68040
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   720
            Width           =   2175
         End
         Begin VB.Frame Frame2 
            Caption         =   "P r e c i o s"
            Height          =   1575
            Index           =   2
            Left            =   -67440
            TabIndex        =   48
            Top             =   2040
            Width           =   1335
            Begin VB.CheckBox Check1 
               Caption         =   "Precio 1"
               Height          =   375
               Index           =   5
               Left            =   120
               TabIndex        =   53
               Top             =   240
               Width           =   1095
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Precio 2"
               Height          =   375
               Index           =   6
               Left            =   120
               TabIndex        =   52
               Top             =   480
               Width           =   1095
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Precio 3"
               Height          =   375
               Index           =   7
               Left            =   120
               TabIndex        =   51
               Top             =   720
               Width           =   1095
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Precio 4"
               Height          =   375
               Index           =   8
               Left            =   120
               TabIndex        =   50
               Top             =   960
               Width           =   1095
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Precio 5"
               Height          =   375
               Index           =   9
               Left            =   120
               TabIndex        =   49
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
            TabIndex        =   47
            Top             =   2880
            Width           =   1815
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
            Index           =   3
            Left            =   2640
            TabIndex        =   99
            Top             =   480
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
            Index           =   4
            Left            =   1800
            TabIndex        =   98
            Top             =   360
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
            Index           =   5
            Left            =   840
            TabIndex        =   97
            Top             =   360
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
            Index           =   7
            Left            =   6600
            TabIndex        =   96
            Top             =   480
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
            Index           =   8
            Left            =   5760
            TabIndex        =   95
            Top             =   360
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
            Index           =   27
            Left            =   4680
            TabIndex        =   94
            Top             =   360
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
            Index           =   33
            Left            =   3840
            TabIndex        =   93
            Top             =   840
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
            Index           =   34
            Left            =   3840
            TabIndex        =   92
            Top             =   1320
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
            Index           =   35
            Left            =   3840
            TabIndex        =   91
            Top             =   1800
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
            Index           =   36
            Left            =   3840
            TabIndex        =   90
            Top             =   2280
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
            Index           =   37
            Left            =   3840
            TabIndex        =   89
            Top             =   2760
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
            Left            =   3840
            TabIndex        =   88
            Top             =   3240
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
            Index           =   6
            Left            =   120
            TabIndex        =   87
            Top             =   960
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
            Index           =   28
            Left            =   120
            TabIndex        =   86
            Top             =   1440
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
            Index           =   29
            Left            =   120
            TabIndex        =   85
            Top             =   1920
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
            Index           =   30
            Left            =   120
            TabIndex        =   84
            Top             =   2400
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
            Index           =   31
            Left            =   120
            TabIndex        =   83
            Top             =   2880
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
            Index           =   32
            Left            =   120
            TabIndex        =   82
            Top             =   3360
            Width           =   600
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
            Index           =   39
            Left            =   -71880
            TabIndex        =   81
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
            Index           =   40
            Left            =   -71880
            TabIndex        =   80
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
            Index           =   41
            Left            =   -71880
            TabIndex        =   79
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
            Index           =   42
            Left            =   -71880
            TabIndex        =   78
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
            Index           =   43
            Left            =   -74880
            TabIndex        =   77
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
            Index           =   44
            Left            =   -74880
            TabIndex        =   76
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
            Index           =   45
            Left            =   -74880
            TabIndex        =   75
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
            Index           =   46
            Left            =   -74880
            TabIndex        =   74
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
            Index           =   47
            Left            =   -71880
            TabIndex        =   73
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
            Index           =   48
            Left            =   -70200
            TabIndex        =   72
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
            Index           =   49
            Left            =   -70200
            TabIndex        =   71
            Top             =   2520
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Polls/Jab"
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
            Index           =   50
            Left            =   -70200
            TabIndex        =   70
            Top             =   1080
            Width           =   1005
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Pollos"
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
            Index           =   51
            Left            =   -70200
            TabIndex        =   69
            Top             =   2040
            Width           =   525
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
            Index           =   52
            Left            =   -70200
            TabIndex        =   68
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
            Index           =   53
            Left            =   -74880
            TabIndex        =   67
            Top             =   2520
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Recepción"
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
            Index           =   54
            Left            =   -68040
            TabIndex        =   66
            Top             =   480
            Width           =   930
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
            Index           =   55
            Left            =   -74880
            TabIndex        =   65
            Top             =   2880
            Width           =   945
         End
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
         Index           =   23
         Left            =   2640
         TabIndex        =   45
         Top             =   480
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
         Index           =   22
         Left            =   1800
         TabIndex        =   44
         Top             =   360
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
         Index           =   21
         Left            =   840
         TabIndex        =   43
         Top             =   360
         Width           =   705
         WordWrap        =   -1  'True
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
         Index           =   12
         Left            =   120
         TabIndex        =   42
         Top             =   840
         Width           =   600
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
         Left            =   6600
         TabIndex        =   41
         Top             =   480
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
         Left            =   5760
         TabIndex        =   40
         Top             =   360
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
         Left            =   4680
         TabIndex        =   39
         Top             =   360
         Width           =   705
         WordWrap        =   -1  'True
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
         Index           =   13
         Left            =   120
         TabIndex        =   38
         Top             =   1320
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
         Index           =   14
         Left            =   120
         TabIndex        =   37
         Top             =   1800
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
         Index           =   15
         Left            =   120
         TabIndex        =   36
         Top             =   2280
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
         Index           =   16
         Left            =   120
         TabIndex        =   35
         Top             =   2760
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
         Index           =   17
         Left            =   120
         TabIndex        =   34
         Top             =   3240
         Width           =   600
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
         Index           =   18
         Left            =   3960
         TabIndex        =   33
         Top             =   840
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
         Index           =   19
         Left            =   3960
         TabIndex        =   32
         Top             =   1320
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
         Index           =   20
         Left            =   3960
         TabIndex        =   31
         Top             =   1800
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
         Index           =   24
         Left            =   3960
         TabIndex        =   30
         Top             =   2280
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
         Index           =   25
         Left            =   3960
         TabIndex        =   29
         Top             =   2760
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
         Index           =   26
         Left            =   3960
         TabIndex        =   28
         Top             =   3240
         Width           =   600
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
      TabIndex        =   10
      Top             =   5040
      Width           =   4635
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IND As Integer
Dim CU As Integer
Dim llave1
Dim cop_llave As rdoResultset
Dim PSCOP_LLAVE As rdoQuery
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
        cont.AddItem tab_mayor!tab_nomlargo & String(15, " ") & tab_mayor!tab_numtab
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub

Public Function PROC_cnt()
Dim tempo As String
Dim NUMCAMPO As Integer
Dim wBusca As String
Dim OJO As String * 1
If PUB_CODTRA = 0 Then
   wBusca = "SELECT * FROM CONTABILIDAD WHERE CNT_CODCIA=?  ORDER BY CNT_CODTRA "
Else
   wBusca = "SELECT * FROM contabilidad WHERE CNT_CODCIA=? AND cnt_CODTRA = ?  ORDER BY CNT_CODTRA "
End If
If UNICO <> wBusca Then
   Set PSX = CN.CreateQuery("", wBusca)
End If
PSX.rdoParameters(0) = LK_CODCIA

If PUB_CODTRA <> 0 Then
  PSX.rdoParameters(1) = PUB_CODTRA
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
   Grid1.Visible = True
   Grid1.Rows = 2
   Grid1.TextMatrix(1, 3) = "No hay registros"
   Screen.MousePointer = 0
   llave1 = ""
   UNICO = wBusca
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
Grid1.ColWidth(1) = 500
Grid1.ColWidth(2) = 300
Grid1.ColWidth(3) = 2400
fila = 0
Grid1.Visible = False
Do Until X.EOF 'Or fila = 50
    fila = fila + 1
    Grid1.TextMatrix(fila, 0) = fila 'Nulo_Valors(X.rdoColumns(3))
    Grid1.TextMatrix(fila, 1) = X.rdoColumns(0)
    Grid1.TextMatrix(fila, 2) = Trim(X.rdoColumns(1))
    Grid1.TextMatrix(fila, 3) = X.rdoColumns(2)
    X.MoveNext
    Grid1.Rows = Grid1.Rows + 1
Loop
Grid1.Visible = True
Grid1.TextMatrix(fila + 1, 3) = "                * * *    END    * * "
End Function



Public Sub GRABAR_DEF()
If Left(cmdModificar.Caption, 2) = "&G" Then
   cnt_llave.Edit
Else
   cnt_llave.AddNew
   cnt_llave(0) = LK_CODCIA
   cnt_llave(1) = FrmCont.TxtDef(0).text
   cnt_llave(2) = FrmCont.TxtDef(1).text
End If

For i = 3 To 38
   If cnt_llave(i).Type = 1 Then
      cnt_llave(i) = Nulo_Valors(FrmCont.TxtDef(i).text)
   Else
      cnt_llave(i) = Val(FrmCont.TxtDef(i).text)
   End If
Next i

cnt_llave!CNT_CODCIA = LK_CODCIA

cnt_llave.Update

End Sub



Private Sub cmdagregar_Click()
'On Error GoTo ESCAPA
If Left(cmdAgregar.Caption, 2) = "&A" Then
    cmdAgregar.Caption = "&Grabar"
    cmdCancelar.Enabled = True
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    DESBLOQUEA_TEXT
    TxtDef(0).Locked = False
    LIMPIA_DEF
    cmdCopiar.Enabled = False
    chetransa.Enabled = False
    cmdconsultar.Enabled = False
    'AGREGAMOS EN BLANCO
Else
     If TxtDef(0).text = "" Or Len(TxtDef(0).text) = 0 Then
          MsgBox "Ingrese , Codigo de Transacción..???", 48, Pub_Titulo
          TxtDef(0).SetFocus
          Exit Sub
      End If
      If TxtDef(1).text = "" Or Len(TxtDef(1).text) = 0 Then
          MsgBox "Ingrese, Secuencia de Transacción..???", 48, Pub_Titulo
          TxtDef(1).SetFocus
          Exit Sub
      End If
'      If EXISTE_DEF(Val(FrmCont.TxtDef(0).text), FrmCont.TxtDef(1).text) Then
'        MENSAJE_DEF "YA Existe Definición Contables .."
'        FrmDef.ListExiste.SetFocus
'       Exit Sub
 '     End If
      
      If CONSIS_CUENTA = 1 Then
          Exit Sub
       End If
      '"SI GRABA.."
      Screen.MousePointer = 11
      GRABAR_DEF
      MENSAJE_DEF "Registro, AGREGADO ... "
      cmdAgregar.Caption = "&Agregar"
      cmdCancelar.Enabled = True
      cmdEliminar.Enabled = True
      cmdModificar.Enabled = True
      BLOQUEA_TEXT
      LIMPIA_DEF
      cmdCancelar.Enabled = True
      TxtDef(0).Locked = False
      TxtDef(0).Enabled = True
      TxtDef(0).SetFocus
      Screen.MousePointer = 0
End If
Exit Sub
    
'ESCAPA:
   
   If Err.Number = 40002 Then
       MsgBox "Hay Error en la LLave ..ESTA DUPLICADO !!! "
   Else
       MsgBox Err.Number & "  " & Err.Description & "   ...  LLAMAR A COMPUTO"
   End If
    
   Exit Sub
    
fin:

End Sub

Private Sub cmdAgregar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    TxtDef(0).SetFocus
End If

End Sub

Private Sub cmdcancelar_Click()
If Left(cmdAgregar.Caption, 2) = "&A" And Left(cmdModificar.Caption, 2) = "&M" Then
    LIMPIA_DEF
    TxtDef(0).Locked = False
    MENSAJE_DEF "Inicializar ... !!!    "
    TxtDef(0).Enabled = True
    TxtDef(0).SetFocus
    cmdCopiar.Enabled = False
    chetransa.Enabled = False
    cmdconsultar.Enabled = False
    Exit Sub
End If
     Screen.MousePointer = 11
     If Left(cmdModificar.Caption, 2) = "&G" Then
        cmdModificar.Caption = "&Modificar"
        LLENA_DEF 1
        cmdCopiar.Enabled = True
        chetransa.Enabled = True
        cmdconsultar.Enabled = True
        TxtDef(0).Locked = True
     Else
        cmdAgregar.Caption = "&Agregar"
        LIMPIA_DEF
        TxtDef(0).Locked = False
        TxtDef(0).SetFocus
     End If
     cmdCerrar.Caption = "&Cerrar"
     cmdAgregar.Enabled = True
     cmdEliminar.Enabled = True
     cmdModificar.Enabled = True
     BLOQUEA_TEXT
     MENSAJE_DEF "Inicializar... !!! "
     Screen.MousePointer = 0

End Sub

Private Sub cmdCerrar_Click()
Unload FrmCont
End Sub

Private Sub cmdCerrar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    TxtDef(0).SetFocus
End If

End Sub

Private Sub cmdConsultar_Click()
Dim i As Integer
Screen.MousePointer = 11
Consultar.Width = 9300 '9495
Consultar.Height = 2775
Consultar.Left = 10
Consultar.Top = 1440
Consultar.RowHeight(0) = 400
SQ_OPER = 2
PUB_CODTRA = Val(TxtDef(0).text)
PUB_CODCIA = LK_CODCIA
LEER_CNT_LLAVE
If cnt_mayor.EOF Then
 Screen.MousePointer = 0
 Exit Sub
End If
fila = 1
Consultar.Rows = 1
Consultar.Cols = cnt_mayor.rdoColumns.count
Do Until cnt_mayor.EOF
 Consultar.Rows = Consultar.Rows + 1
 For i = 0 To cnt_mayor.rdoColumns.count - 1
  Consultar.TextMatrix(0, i) = Mid(cnt_mayor.rdoColumns(i).Name, 5, Len(cnt_mayor.rdoColumns(i).Name))
  If Not IsNull(cnt_mayor.rdoColumns(i)) Then
    If IsNumeric(cnt_mayor.rdoColumns(i)) Then
     Consultar.TextMatrix(fila, i) = Val(cnt_mayor.rdoColumns(i))
    Else
     Consultar.TextMatrix(fila, i) = cnt_mayor.rdoColumns(i)
    End If
  End If
  Consultar.ColWidth(i) = 700
 Next i
 fila = fila + 1
 cnt_mayor.MoveNext
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
pub_cadena = "SELECT * FROM DEFCONT "
Set otro_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
otro_llave.Requery
 
otro_llave.AddNew
For fila = 0 To SUT_LLAVE.rdoColumns.count - 1
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
    FrmDef.TxtDef(0).text = ""
    FrmDef.TxtDef(0).Locked = False
    LIMPIA_DEF
    MENSAJE_DEF "Registro, ELIMINADO ... "
    Screen.MousePointer = 0
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
If CONSIS_CUENTA = 1 Then
   MENSAJE_DEF "Alguna cuenta no existe..."
   Exit Sub
End If

If Left(cmdModificar.Caption, 2) = "&M" Then
    cmdModificar.Caption = "&Grabar"
    cmdEliminar.Enabled = False
    cmdAgregar.Enabled = False
    cmdCancelar.Enabled = True
    DESBLOQUEA_TEXT
    TxtDef(0).Locked = True
    TxtDef(1).Locked = True
    TxtDef(2).Locked = True
    cmdCopiar.Enabled = False
    chetransa.Enabled = False
    cmdconsultar.Enabled = False
    FrmCont.TxtDef(3).SetFocus
Else

    Screen.MousePointer = 11
    GRABAR_DEF
    MENSAJE_DEF "Registro,MODIFICADO... "
    cmdModificar.Caption = "&Modificar"
    cmdCancelar.Enabled = True
    cmdEliminar.Enabled = True
    cmdAgregar.Enabled = True
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

Private Sub CUENTAS_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 FCUENTAS.Visible = False
 TxtDef(IND).SetFocus
 Exit Sub
End If
If KeyAscii = 13 Then
  If Left(CUENTAS.text, 1) = " " Then
    TxtDef(IND).SetFocus
    Exit Sub
  End If
  TxtDef(IND).text = Left(CUENTAS.text, 12)
  FCUENTAS.Visible = False
  TxtDef(IND).SetFocus
End If
End Sub

Private Sub CUENTAS_LostFocus()
FCUENTAS.Visible = False
TxtDef(IND).SetFocus
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
LLENADOS FrmDef.def_abreviado, 13
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
PUB_CODCIA = "00"
FrmDatArti.Caption = "Def_art_gru.  -  TAB_TIPREG = " & PUB_TIPREG
Load FrmDatArti
FrmDatArti.Show 1
DoEvents
PUB_CODCIA = LK_CODCIA
LLENADOS FrmDef.def_art_gru, 13
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
LLENADOS FrmDef.def_tipdoc, 8
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
LLENADOS FrmDef.def_tipmov, 4
def_tipmov.SetFocus
SendKeys "%{up}"

End Sub

Private Sub def_tipmov_ref_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos
If KeyCode <> 45 Then
 Exit Sub
End If
wpos = def_tipmov_ref.ListIndex
PUB_TIPREG = Mid(def_tipmov_ref.ToolTipText, 13, Len(def_tipmov_ref.ToolTipText))
Load FrmDatArti
PUB_CODCIA = "00"
FrmDatArti.Caption = "TIP-MOVIMIENTO -  TAB_TIPREG = " & PUB_TIPREG
Load FrmDatArti
FrmDatArti.Show 1
DoEvents
PUB_CODCIA = "00"
LLENADOS FrmDef.def_tipmov_ref, 4
def_tipmov_ref.SetFocus
SendKeys "%{up}"

End Sub

Private Sub Form_Load()
'MSTDEF.DSN = PUB_DSN
cade = "SELECT * FROM COPARAM WHERE COP_CODCIA = ?"
Set PSCOP_LLAVE = CN.CreateQuery("", cade)
Set cop_llave = PSCOP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
PSCOP_LLAVE.rdoParameters(0) = LK_CODCIA

LLENA_COMAEST

CU = 0
BLOQUEA_TEXT


LIMPIA_DEF
F1.Visible = True

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
End Sub
Public Sub DESBLOQUEA_TEXT()
For i = 1 To 38
    TxtDef(i).Enabled = True
Next i
End Sub

Public Sub LIMPIA_DEF()
Dim i As Integer
For i = 0 To 38
    TxtDef(i).text = ""
Next i
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
    PUB_CODTRA = Val(Grid1.TextMatrix(Grid1.Row, 2))
   
    MSTDEF.Visible = False
    LLENA_DEF 0
    cmdCancelar.Enabled = True
    TxtDef(0).Locked = True
    TxtDef(2).Locked = True
    cmdCopiar.Enabled = True
    chetransa.Enabled = True
    cmdconsultar.Enabled = True
    cmdModificar.SetFocus
End If


End Sub

Private Sub ListExiste_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  'sn_mensaje = " ¿Desea aun Grabar estos datos del  Cliente ... ? "
  'pub_respuesta = msgbox(sn_mensaje, pub_estilo, pub_titulo)
  'If pub_respuesta = vbYes Then   ' El usuario eligió
  '   PASA = True
  '   frmCLI.ListExiste.Visible = False
  '   cmdAgregar_Click
  '   KeyAscii = 0
  'Else
   ' PASA = False
    FrmDef.ListExiste.Visible = False
    FrmDef.TxtDef(1).SetFocus
    KeyAscii = 0
  End If

End Sub

Private Sub ListExiste_LostFocus()
 FrmDef.ListExiste.Visible = False
 FrmDef.TxtDef(1).SetFocus
 
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
Select Case Index
Case 2
    KeyAscii = tempo
Case 0
   If Not car < "0" And car > "9" Then
      If KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
        Beep
        Exit Sub
       End If
   End If
   If KeyAscii = 13 Then
        If Trim(TxtDef(0).text) = "" Then
          Exit Sub
        End If
        If FrmCont.cmdAgregar.Enabled = False Or FrmCont.cmdModificar.Enabled = False Then
           Exit Sub
        End If
        IND = Index
        llave1 = ""
        UNICO = ""
        PUB_CODTRA = Val(TxtDef(0).text)
        KeyAscii = 0
        MSTDEF.Visible = True
        Call PROC_cnt
        Grid1.Col = 3
        Grid1.Row = 1
        Grid1.SetFocus
        Exit Sub
    End If
Case 1
   If Not car < "0" And car > "9" Then
      If KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
        Beep
        Exit Sub
       End If
   End If

Case 4, 7, 10, 13, 16, 19, 22, 25, 28, 31, 34, 37
    If KeyAscii <> 68 And KeyAscii <> 72 And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 32 Then
        KeyAscii = 0
        Beep
        Exit Sub
    End If

Case 3, 6, 9, 12, 15, 18, 21, 24, 27, 30, 33, 36
    If KeyAscii = 13 Then
      IND = Index
      FCUENTAS.Visible = True
      CUENTAS.SetFocus
      Exit Sub
    End If

End Select

End Sub

Public Sub LLENA_DEF(ban As Integer)
Dim i As Integer
If ban = 0 Then
    SQ_OPER = 1
    PUB_CODTRA = Grid1.TextMatrix(Grid1.Row, 2)
    PUB_SECUENCIA = Grid1.TextMatrix(Grid1.Row, 3)
    PUB_CODCIA = LK_CODCIA
    LEER_CNT_LLAVE
    PUB_CODTRA = Grid1.TextMatrix(Grid1.Row, 2)
    PUB_SECUENCIA = Grid1.TextMatrix(Grid1.Row, 3)
    LEER_SUT_LLAVE
    If SUT_LLAVE.EOF Then
       MsgBox "Revisar en sub-Transacciones... "
       Exit Sub
    End If
    
End If
For i = 3 To 38
    FrmCont.TxtDef(i).text = Trim(Nulo_Valors(cnt_llave.rdoColumns(i)))
Next i
FrmCont.TxtDef(2).text = SUT_LLAVE!SUT_DESCRIPCION
FrmCont.TxtDef(0).text = cnt_llave.rdoColumns(1)
FrmCont.TxtDef(1).text = cnt_llave.rdoColumns(2)

End Sub




Public Sub LLENA_COMAEST()
Dim CO As rdoResultset
Dim cade As String
Dim wcuenta As String * 12
Dim wdescri As String * 30
cop_llave.Requery
If cop_llave.EOF Then
  CUENTAS.AddItem " No Existen en CIA"
  Exit Sub
End If
cade = "SELECT * FROM COMAEST WHERE COM_CODCIA = '" & LK_CODCIA & "' AND COM_NIVEL = " & Nulo_Valor0(cop_llave!COP_NIVEL_AFECTACION) & " ORDER BY COM_CUENTA"
Set CO = CN.OpenResultset(cade, rdOpenKeyset, rdConcurValues)
CO.Requery
CUENTAS.Clear
If CO.EOF Then
 CUENTAS.AddItem " No Existen en CIA"
 Exit Sub
End If
Do Until CO.EOF
  wcuenta = Trim(CO!com_cuenta)
  wdescri = Trim(CO!COM_DESCRIPCION)
  CUENTAS.AddItem wcuenta & " > " & wdescri
  CO.MoveNext
Loop
CUENTAS.ListIndex = 0
End Sub

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
