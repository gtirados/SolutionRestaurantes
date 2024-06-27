VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Gestion Contable S.A.C."
   ClientHeight    =   6510
   ClientLeft      =   -210
   ClientTop       =   735
   ClientWidth     =   9480
   Icon            =   "MDIGeneral.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "MDIGeneral.frx":0442
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   820
      ButtonWidth     =   714
      ButtonHeight    =   661
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   61
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "FORMGEN"
            Object.ToolTipText     =   "Transacciones"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Tablas"
            Object.ToolTipText     =   "Mant. Tablas"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "ARTI"
            Object.ToolTipText     =   "Articulos"
            Object.Tag             =   ""
            ImageIndex      =   65
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "PRECIOS"
            Object.ToolTipText     =   "Precios"
            Object.Tag             =   ""
            ImageIndex      =   66
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "CLIENTES"
            Object.ToolTipText     =   "Clientes"
            Object.Tag             =   ""
            ImageIndex      =   78
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "BANCOS"
            Object.ToolTipText     =   "Bancos"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "CUENTAS"
            Object.ToolTipText     =   "Estado Cuentas"
            Object.Tag             =   ""
            ImageIndex      =   77
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "CHEQUES"
            Object.ToolTipText     =   "Cheques"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "KARDEX"
            Object.ToolTipText     =   "Kardex"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Password"
            Object.ToolTipText     =   "Password"
            Object.Tag             =   ""
            ImageIndex      =   46
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "ALTERNO"
            Object.ToolTipText     =   "Modo Alterno"
            Object.Tag             =   ""
            ImageIndex      =   56
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Excel"
            Object.ToolTipText     =   "Excel"
            Object.Tag             =   ""
            ImageIndex      =   71
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button23 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button24 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button25 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button26 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button27 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button28 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button29 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button30 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button31 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button32 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button33 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button34 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button35 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button36 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button37 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button38 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button39 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button40 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button41 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button42 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button43 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button44 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button45 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button46 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button47 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button48 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button49 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button50 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button51 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button52 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button53 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button54 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button55 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button56 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button57 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button58 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button59 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button60 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button61 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CommandButton cmdmes 
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdmes 
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdmes 
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdmes 
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdmes 
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdmes 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdmes 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdmes 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdmes 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdmes 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdmes 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdmes 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdmes 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdmes 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdmes 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdmes 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox tperiodo 
         BackColor       =   &H00FFFFFF&
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
         Height          =   405
         Left            =   6000
         TabIndex        =   6
         Top             =   0
         Width           =   2895
      End
      Begin ComctlLib.ProgressBar ProBar 
         Height          =   195
         Left            =   9000
         TabIndex        =   3
         Top             =   195
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.TextBox txtderechos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
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
         Left            =   8985
         TabIndex        =   4
         Text            =   "Cargando Accesos..."
         Top             =   10
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton TXTCIA 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   2760
      End
      Begin VB.PictureBox Document1 
         Height          =   225
         Left            =   11400
         ScaleHeight     =   165
         ScaleWidth      =   195
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   120
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport TablasC 
      Left            =   120
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\Archivos de programa\PROYECTO WINDOWS NT\tablas.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   6240
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   7
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1429
            MinWidth        =   1411
            TextSave        =   "MAYÚS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NÚM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1587
            TextSave        =   "09:35 a.m."
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   6174
            MinWidth        =   6174
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4339
            MinWidth        =   4339
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   4210752
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   79
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":6FFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":7110
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":742A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":7744
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":7A5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":17A40
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":17D5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":18074
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1838E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":184A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":187BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":18AD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":18DEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":19108
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1921A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1932C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1943E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":19550
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":19662
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":19774
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":19886
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":19998
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":19AAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":19BBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":19CCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":19E68
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1A002
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1A224
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1A45A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1B338
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1B4D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1B66C
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1B766
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1B860
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1B95A
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1BA54
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1BCE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1C378
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1C472
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1C60C
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1C7A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1C940
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1CADA
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1CC74
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1CE0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1D640
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1D7DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1D974
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1DA2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1DADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1DBEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1DD00
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1DE12
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1DF24
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1E036
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1E868
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1EA02
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1EB9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1EECE
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":1F068
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":2725A
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":273F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":2758E
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":278A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":27BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":27EDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":281F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":28510
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":2882A
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":28B44
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":28E5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":29178
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":29492
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":297AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":29AC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":29DE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":2A0FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":2A414
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIGeneral.frx":2A72E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu menuTit1 
      Caption         =   "&Archivos"
      Enabled         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Articulos"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^A
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Centro de Costos"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^I
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Ven&dedores"
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   ^D
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Bancos"
         Enabled         =   0   'False
         Index           =   3
         Shortcut        =   ^B
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "-"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Tablas del Sistema"
         Enabled         =   0   'False
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "C&ompañia"
         Enabled         =   0   'False
         Index           =   6
         Shortcut        =   ^F
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Empresa &General"
         Enabled         =   0   'False
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Tablas"
         Enabled         =   0   'False
         Index           =   8
         Shortcut        =   ^T
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Definición de Transacciones"
         Enabled         =   0   'False
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Calendario de Sistema /T. Cambio"
         Enabled         =   0   'False
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "-"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "C&ambiar Clave..."
         Enabled         =   0   'False
         Index           =   12
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Reiniciar  &Usuario"
         Enabled         =   0   'False
         Index           =   13
         Shortcut        =   ^U
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Reiniciar Sistema"
         Enabled         =   0   'False
         Index           =   14
         Shortcut        =   ^R
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Salir"
         Enabled         =   0   'False
         Index           =   15
         Shortcut        =   +^{F9}
      End
   End
   Begin VB.Menu menuTit2 
      Caption         =   "&Tiempo Real"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu SubmenuTit2 
         Caption         =   "&Transacciones"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^G
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "C&osteo de Productos"
         Index           =   1
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "&Cierre del Dia"
         Enabled         =   0   'False
         Index           =   3
      End
   End
   Begin VB.Menu menutit3 
      Caption         =   "&Contabilidad"
      Enabled         =   0   'False
      Begin VB.Menu submenutit3 
         Caption         =   "Definición de &Estructura"
         Index           =   0
      End
      Begin VB.Menu submenutit3 
         Caption         =   "&Plan Contable"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu submenutit3 
         Caption         =   "&Plantilla para Vouchers"
         Index           =   2
      End
      Begin VB.Menu submenutit3 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu submenutit3 
         Caption         =   "&Gestión de Vouchers"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu submenutit3 
         Caption         =   "D&iario General"
         Index           =   5
      End
      Begin VB.Menu submenutit3 
         Caption         =   "&Procesar Libros Contables"
         Index           =   6
      End
      Begin VB.Menu submenutit3 
         Caption         =   "&Ajuste Mensual por Dif. Tipo Cambio"
         Index           =   7
      End
      Begin VB.Menu submenutit3 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu submenutit3 
         Caption         =   "Con&sultas de Cuentas"
         Index           =   9
      End
      Begin VB.Menu submenutit3 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu submenutit3 
         Caption         =   "Pase a Con&tabilidad"
         Enabled         =   0   'False
         Index           =   11
      End
      Begin VB.Menu submenutit3 
         Caption         =   "Procesar C&uentas Destinos"
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu submenutit3 
         Caption         =   "Procesar &Mayorizacion "
         Enabled         =   0   'False
         Index           =   13
      End
      Begin VB.Menu submenutit3 
         Caption         =   "Cierre &Del Periodo de Trabajo "
         Enabled         =   0   'False
         Index           =   14
      End
      Begin VB.Menu submenutit3 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu submenutit3 
         Caption         =   "&Importar Datos"
         Index           =   16
      End
   End
   Begin VB.Menu menutit4 
      Caption         =   "&Listados"
      Enabled         =   0   'False
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Ventas"
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
         Begin VB.Menu menuVent 
            Caption         =   " "
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Compras"
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
         Begin VB.Menu Menucomp 
            Caption         =   " "
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Principales"
         Enabled         =   0   'False
         Index           =   4
         Visible         =   0   'False
         Begin VB.Menu menuAlm 
            Caption         =   " "
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Inventarios"
         Enabled         =   0   'False
         Index           =   6
         Visible         =   0   'False
         Begin VB.Menu menuMoli 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "C&ontabilidad"
         Enabled         =   0   'False
         Index           =   8
         Visible         =   0   'False
         Begin VB.Menu MenuContab 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Disponible 1 "
         Index           =   10
         Visible         =   0   'False
         Begin VB.Menu menudis1 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Disponible 2"
         Index           =   12
         Visible         =   0   'False
         Begin VB.Menu menudis2 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   13
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Disponible 3"
         Index           =   14
         Visible         =   0   'False
         Begin VB.Menu menudis3 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   15
      End
   End
   Begin VB.Menu menutit5 
      Caption         =   "&Utilidades"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu submenutit5 
         Caption         =   "&Backups            "
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu submenutit5 
         Caption         =   "&Mostrar ultimo Reporte"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu submenutit5 
         Caption         =   "&Configuración de Reporte"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu submenutit5 
         Caption         =   "&Grupo de Trabajso"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu submenutit5 
         Caption         =   "&Reporte de Transacciones"
         Enabled         =   0   'False
         Index           =   4
      End
   End
   Begin VB.Menu menutit6 
      Caption         =   "&Herramientas"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu SubmenuTit6 
         Caption         =   "&Consultas"
         Enabled         =   0   'False
         Index           =   0
         Begin VB.Menu menuctacte 
            Caption         =   "Cu&entas Corrientes"
            Shortcut        =   ^E
         End
         Begin VB.Menu menuch 
            Caption         =   "Ctas. Ctes de &Bancos"
            Shortcut        =   ^H
         End
         Begin VB.Menu menuKa 
            Caption         =   "&Kardex"
            Shortcut        =   ^K
         End
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "Consulta de &Documentos"
         Index           =   1
         Shortcut        =   ^N
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "&Actualizar Articulos"
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   ^P
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "&Ventas en Tiempo Real"
         Index           =   4
         Shortcut        =   ^S
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   " &Orden de Compra"
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "-"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   " &Transferencia de Pedidos"
         Enabled         =   0   'False
         Index           =   8
         Visible         =   0   'False
      End
   End
   Begin VB.Menu menuAyuda 
      Caption         =   "&Información"
      Begin VB.Menu menuAyudaSistema 
         Caption         =   "&Acerca de. . .."
      End
      Begin VB.Menu menuRepoAyuda 
         Caption         =   "&Ayuda . . . "
      End
   End
   Begin VB.Menu Menucompania 
      Caption         =   "compania"
      Visible         =   0   'False
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   2
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   3
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   4
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   5
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   6
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   7
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   8
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   9
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   10
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   11
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   12
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   13
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   14
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   15
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   16
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   17
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   18
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   19
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   20
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   21
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   22
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   23
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   24
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   25
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   26
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   27
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   28
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   29
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   30
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   31
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   32
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   33
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   34
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   35
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   36
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   37
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   38
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   39
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   40
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WF_REINICIAR As Integer
Dim Wfile As String
Dim WFORM As String

Public Sub LLENA_ACCESO_OTROS(WSOTROS As String)
Dim W1 As String * 2
Dim i, wPosF, WPosV, cuenta As Integer
Dim sal As Boolean
Dim cade As String
Dim WNUM As Integer
Dim f As Integer
Dim a As Integer
WNUM = 0
wPosF = 0
WPosV = 0
cuenta = 0
WPosV = Len(WSOTROS)
cade = Trim(WSOTROS)
cuenta = 0
wPosF = 1
a = 0
For i = 1 To Len(cade)
 If Mid(cade, i, 1) = "." Then
   a = a + 1
 End If
Next i
lk_OTROS_Count = a
ReDim lk_OTROS(a)
Do Until cuenta = a
   cuenta = cuenta + 1
   wPosF = InStr(wPosF, cade, ".", 1) + 1
   WNUM = Mid(cade, wPosF, 2)
   If Right(WNUM, 1) = "." Then
     WNUM = Left(WNUM, 2)
     wPosF = wPosF - 1
   End If
   lk_OTROS(cuenta) = Format(CStr(WNUM), "00")
Loop

End Sub

Private Sub cmdmes_Click(Index As Integer)
Dim wfecha1
Dim wfecha2
WDIAS = dias_mes(Index, Format(LK_FECHA_COP1, "yyyy"))
wmes = Index
If wmes = 0 Then
 wfecha1 = "01" & "/" & Format(wmes + 1, "00") & "/" & Format(LK_FECHA_COP1, "yyyy")
 wfecha2 = Format(1, "00") & "/" & Format(wmes + 1, "00") & "/" & Format(LK_FECHA_COP1, "yyyy")
ElseIf wmes = 13 Then
 wfecha1 = "28" & "/" & Format(wmes - 1, "00") & "/" & Format(LK_FECHA_COP1, "yyyy")
 wfecha2 = Format(28, "00") & "/" & Format(wmes - 1, "00") & "/" & Format(LK_FECHA_COP1, "yyyy")
ElseIf wmes = 14 Then
 wfecha1 = "29" & "/" & Format(wmes - 2, "00") & "/" & Format(LK_FECHA_COP1, "yyyy")
 wfecha2 = Format(29, "00") & "/" & Format(wmes - 2, "00") & "/" & Format(LK_FECHA_COP1, "yyyy")
ElseIf wmes = 15 Then
 wfecha1 = "30" & "/" & Format(wmes - 3, "00") & "/" & Format(LK_FECHA_COP1, "yyyy")
 wfecha2 = Format(30, "00") & "/" & Format(wmes - 3, "00") & "/" & Format(LK_FECHA_COP1, "yyyy")
Else
 wfecha1 = "01" & "/" & Format(Index, "00") & "/" & Format(LK_FECHA_COP1, "yyyy")
 wfecha2 = Format(WDIAS, "00") & "/" & Format(Index, "00") & "/" & Format(LK_FECHA_COP1, "yyyy")
End If

LK_NRO_MES = wmes
LK_FECHA_COP1 = wfecha1
LK_FECHA_COP2 = wfecha2
If LK_NRO_MES = 0 Then
   MDIForm1.tperiodo.Text = "Pdo.: " & "APERTURA - " & Format(LK_FECHA_COP1, "yyyy")
ElseIf LK_NRO_MES = 13 Then
   MDIForm1.tperiodo.Text = "Pdo.: " & "Ajuste x Inflación - " & Format(LK_FECHA_COP1, "yyyy")
ElseIf LK_NRO_MES = 14 Then
   MDIForm1.tperiodo.Text = "Pdo.: " & "Cierre de Ctas. I/S - " & Format(LK_FECHA_COP1, "yyyy")
ElseIf LK_NRO_MES = 15 Then
   MDIForm1.tperiodo.Text = "Pdo.: " & "Cierre Anual - " & Format(LK_FECHA_COP1, "yyyy")
Else
   MDIForm1.tperiodo.Text = "Periodo: " & Format(LK_FECHA_COP1, "mmmm yyyy")
End If
pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER>=? AND COV_FECHA_VOUCHER <=? AND COV_NRO_MES = " & LK_NRO_MES & "  ORDER BY COV_NRO_VOUCHER, COV_NRO_MOV"
Set PSCOV_VOUCHER = CN.CreateQuery("", pub_cadena)
PSCOV_VOUCHER(0) = 0
PSCOV_VOUCHER(1) = LK_FECHA_DIA
PSCOV_VOUCHER(2) = LK_FECHA_DIA
Set cov_voucher = PSCOV_VOUCHER.OpenResultset(rdOpenKeyset, rdConcurValues)
Dim iFormCount As Integer
On Error GoTo SIGUE
Screen.MousePointer = 11
If Forms.Count - 1 > 0 Then
   For iFormCount = Forms.Count - 1 To 1 Step -1
    If iFormCount <> 1 Then
        Unload Forms(iFormCount)
    End If
   Next iFormCount
End If
Screen.MousePointer = 0
Exit Sub
SIGUE:
Screen.MousePointer = 0
End Sub

Private Sub MDIForm_Load()
Dim WUSU_CIAS As String
Dim WUSU_OTROS As String
Dim h, a, b, C, d, E As Integer
Dim wBAN As Integer
Dim WNUM, wPosF, WPosV, cuenta, cade
Dim Wflag As String * 1
Dim f, g, fg As Integer

'On Error GoTo SALE
Wflag = ""
TXTCIA.Enabled = False
PS_GEN(0) = 0
GEN.Requery
Pub_Titulo = Trim(GEN!GEN_NOMBRE)
PUB_RUTA_REPORTE = Trim(Nulo_Valors(GEN!gen_ruta_reportes))
PUB_RUTA_OTRO = Trim(Nulo_Valors(GEN!gen_ruta_otros))
LK_IGV = Nulo_Valor0(GEN!GEN_IGV)
LK_TIPO_CAMBIO = Nulo_Valor0(GEN!gen_tipo_cambio)
LK_DIG_RUC = Nulo_Valor0(GEN!GEN_DIG_RUC)
CONS_ADMIN = Mid(PUB_RUTA_REPORTE, 1, Len(PUB_RUTA_REPORTE) - 6)

TXTCIA.Visible = False
DoEvents
Splash.Hide
DoEvents
MDIForm1.Show
DoEvents
NOTPermisos
PASSFORM.Show 1
DoEvents
Screen.MousePointer = 11
ProBar.Value = 0
ProBar.Min = 0
ProBar.Max = 15
txtderechos.Visible = True
DoEvents
ProBar.Visible = True
DoEvents
If LK_CODCIA = "" Or LK_CODUSU = "" Then
  MsgBox "Existe un Problema con su Acceso , consulte al Administrador ", 48, Pub_Titulo
  GoTo fin_de_pro
End If

StatusBar1.Enabled = False
PS_PAR(0) = " "
par.Requery
If par.EOF Then
  GoTo PASAR
End If
'If LK_CODUSU <> "ADMIN" Then
' MDIForm1.submenutit5(2).Visible = False
 'MDIForm1.submenutit5(3).Visible = False
'Else
' MDIForm1.submenutit5(2).Visible = True
' MDIForm1.submenutit5(3).Visible = True
'End If
usu.Requery
ProBar.Value = ProBar.Value + 1
Do Until usu.EOF
If Trim(usu!usu_key) = LK_CODUSU Then
   WUSU_CIAS = Trim(Nulo_Valors(usu!USU_CIAS))
   WUSU_OTROS = Trim(Nulo_Valors(usu!USU_OTROS))
   LK_PRECIO = Trim(Nulo_Valors(usu!USU_PRECIO))
   
   Exit Do
End If
usu.MoveNext
Loop
If Len(WUSU_CIAS) = 0 Then
  MsgBox "Usted NO Tiene accesos a ninguna compañia , consulte al Administrador ", 48, Pub_Titulo
  GoTo fin_de_pro
End If

If WUSU_OTROS = "" Then
 lk_OTROS_Count = 0
 ReDim lk_OTROS(0)
Else
 LLENA_ACCESO_OTROS WUSU_OTROS
End If
ProBar.Value = ProBar.Value + 1
WNUM = 0
wPosF = 0
WPosV = 0
cuenta = 0
WPosV = Len(WUSU_CIAS)
cade = Trim(WUSU_CIAS)
cuenta = 0
wPosF = 1
a = 0
ProBar.Value = ProBar.Value + 1
For fila = 1 To Len(cade)
    If Mid(cade, fila, 1) = "." Then
        a = a + 1
    End If
Next fila
ProBar.Value = ProBar.Value + 1
wBAN = -1
h = 0
Do Until cuenta = a
    h = h + 1
    cuenta = cuenta + 1
    DoEvents
    wPosF = InStr(wPosF, cade, ".", 1) + 1
    DoEvents
    WNUM = Mid(cade, wPosF, 2)
    If Right(WNUM, 1) = "." Then
        WNUM = Left(WNUM, 2)
        wPosF = wPosF - 1
    End If
    SQ_OPER = 1
    PUB_CODCIA = Format(WNUM, "00")
    LEER_PAR_LLAVE
    If Not par_llave.EOF Then
       MDIForm1.menuCia(cuenta - 1).Caption = PUB_CODCIA & "-" & Trim(par_llave!PAR_NOMBRE)
       MDIForm1.menuCia(cuenta - 1).Visible = True
       If LK_CODCIA = par_llave!PAR_CODCIA Then wBAN = h - 1
    Else
       If a = 1 Then ' Cuando la unica compañia de acceso no exite
         MsgBox "Su Compañia de Acceso NO Existe , consulte al Administrador, Procede a Salir del Sistema", 48, Pub_Titulo
         GoTo fin_de_pro
       End If
       MDIForm1.menuCia(cuenta - 1).Caption = "* " & PUB_CODCIA & " , NO Procede"
       MDIForm1.menuCia(cuenta - 1).Visible = True
    End If
Loop
ProBar.Value = ProBar.Value + 1
If wBAN = -1 Then
  For a = 0 To h - 1
    If Left(MDIForm1.menuCia(a).Caption, 1) <> "*" Then
        wBAN = a
        Exit For
    End If
  Next a
  If wBAN = -1 Then
    MsgBox "Su(s) Compañias de Acceso NO Existen , consulte al Administrador, Procede a Salir del Sistema", 48, Pub_Titulo
    GoTo fin_de_pro
  End If
End If

For a = h To 40
  MDIForm1.menuCia(a).Visible = False
Next a
MDIForm1.menuCia(wBAN).Checked = True
'asegura la compañia por defecto
ACTUALIZA_CIA Left(MDIForm1.menuCia(wBAN).Caption, 2)
ProBar.Value = ProBar.Value + 1
PASAR:
PS_GEN(0) = 0
GEN.Requery
ProBar.Value = ProBar.Value + 1



MDIForm1.StatusBar1.Panels(5).Text = Trim(GEN!GEN_NOMBRE)
MDIForm1.StatusBar1.Panels(6).Text = LK_CODUSU
MDIForm1.StatusBar1.Panels(4).Text = "Fecha de la Empresa :  " & Format(LK_FECHA_DIA, "dd, mmmm yyyy")
MDIForm1.StatusBar1.Panels(7).Text = MDIForm1.menuCia(wBAN).Caption
ProBar.Value = ProBar.Value + 1
MDIForm1.TXTCIA.Caption = MDIForm1.menuCia(wBAN).Caption
CONST_SERVER = Trim(Nulo_Valors(GEN!GEN_SERVER))
CONST_UID = Trim(Nulo_Valors(GEN!GEN_UID))
CONST_PWD = Trim(Nulo_Valors(GEN!GEN_PWD))
ProBar.Value = ProBar.Value + 1
If Len(CONST_SERVER) = 0 Or Len(CONST_UID) = 0 Then
  MsgBox "ADVERTENCIA ..!! DEBES DEFINIR EL NOMBRE DEL SERVIDOR , UID , PWD", 48, Pub_Titulo
End If
PUB_ODBC = "ODBC;DRIVER={SQL Server};UID=" & CONST_UID & ";PWD=" & CONST_PWD & ";APP=Data From Designer;WSID=" & CONST_SERVER & ";LANGUAGE=us_english;DATABASE=BDATOS"

SQ_OPER = 2
PUB_TIPREG = 99
PUB_CODCIA = "00"
LEER_TAB_LLAVE
If tab_mayor.EOF Then
 MsgBox " No existe TIPREG 99 en tablas crearlas con titulos del Menu correspondiente..", 48, Pub_Titulo
End If
fila = 0
SQ_OPER = 0
Do Until tab_mayor.EOF
If UCase(Trim(tab_mayor!tab_nomlargo)) = "CONTABILIDAD" Then
'Else
'GoTo O_CON
    MDIForm1.SubmenuTit4(fila).Caption = "&" + Trim(tab_mayor!tab_nomlargo)
    If Val(tab_mayor!tab_nomcorto) = 1 Then
     MDIForm1.SubmenuTit4(fila).Visible = True
     MDIForm1.SubmenuTit4(fila + 1).Visible = True
     SQ_OPER = 9
    End If
 End If
 fila = fila + 2
O_CON:
 tab_mayor.MoveNext
Loop
If SQ_OPER = 9 Then
 MDIForm1.SubmenuTit4(15).Visible = False
End If





' * Permisos para los Menus *
Screen.MousePointer = 11
Permisos
ProBar.Value = ProBar.Value + 1
Screen.MousePointer = 11
' * Carga menu de Reporte *
SQ_OPER = 2
PUB_INICIO = 8100
LEER_TRA_LLAVE
ProBar.Value = ProBar.Value + 1
a = 0
b = 0
C = 0
d = 0
E = 0
f = 0
g = 0
fg = 0
Do Until tra_menu.EOF
 If tra_menu!TRA_FLAG_EXT <> "A" Then
   GoTo OTRO_MENU
 End If
 ' CONDICION SOLO PARA MENU DE CONTABILIDAD
 If tra_menu(0) >= 8900 And tra_menu(0) < 9099 Then
 Else
    GoTo OTRO_MENU
 End If
 

 If tra_menu(0) >= 8100 And tra_menu(0) < 8299 Then
     If a > 0 Then
        Load MDIForm1.menuVent(a)
     End If
    MDIForm1.menuVent(a).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    a = a + 1
 ElseIf tra_menu(0) >= 8300 And tra_menu(0) < 8499 Then
    If b > 0 Then
        Load MDIForm1.Menucomp(b)
    End If
    MDIForm1.Menucomp(b).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    b = b + 1
 ElseIf tra_menu(0) >= 8500 And tra_menu(0) < 8699 Then
    If C > 0 Then
        Load MDIForm1.menuAlm(C)
    End If
    MDIForm1.menuAlm(C).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    C = C + 1
 ElseIf tra_menu(0) >= 8700 And tra_menu(0) < 8899 Then
    If d > 0 Then
        Load MDIForm1.menuMoli(d)
    End If
    MDIForm1.menuMoli(d).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    d = d + 1
 ElseIf tra_menu(0) >= 8900 And tra_menu(0) < 9099 Then
    If E > 0 Then
        Load MDIForm1.MenuContab(E)
    End If
    MDIForm1.MenuContab(E).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    E = E + 1
 ElseIf tra_menu(0) >= 9100 And tra_menu(0) < 9299 Then
    If f > 0 Then
        Load MDIForm1.menudis1(f)
    End If
    MDIForm1.menudis1(f).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    f = f + 1
 ElseIf tra_menu(0) >= 9300 And tra_menu(0) < 9499 Then
    If g > 0 Then
        Load MDIForm1.menudis2(g)
    End If
    MDIForm1.menudis2(g).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    g = g + 1
 ElseIf tra_menu(0) >= 9500 And tra_menu(0) < 9699 Then
    If fg > 0 Then
        Load MDIForm1.menudis3(fg)
    End If
    MDIForm1.menudis3(fg).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    fg = fg + 1
 End If

OTRO_MENU:
 tra_menu.MoveNext
Loop


' QUITAR EL ACCESO MENOS A CONTABILIDAD
SQ_OPER = 2
PUB_TIPREG = 99
PUB_CODCIA = "00"
LEER_TAB_LLAVE
If tab_mayor.EOF Then
 MsgBox " No existe TIPREG 99 en tablas crearlas con titulos del Menu correspondiente..", 48, Pub_Titulo
End If
fila = 0
SQ_OPER = 0
Do Until tab_mayor.EOF
    If UCase(Trim(tab_mayor!tab_nomlargo)) = "CONTABILIDAD" Then
    Else
    ' MDIForm1.SubmenuTit4(fila).Caption = "&" + Trim(tab_mayor!tab_nomlargo)
     If Val(tab_mayor!tab_nomcorto) = 1 Then
      MDIForm1.SubmenuTit4(fila).Visible = False
'      MDIForm1.SubmenuTit4(fila + 1).Visible = False
      SQ_OPER = 9
     End If
    End If
    fila = fila + 2
    tab_mayor.MoveNext
Loop
If SQ_OPER = 9 Then
 'MDIForm1.SubmenuTit4(15).Visible = False
End If



If LK_CODUSU <> "ADMIN" Then
 MDIForm1.SubmenuTit1(8).Caption = "&Usuarios"
 MDIForm1.Toolbar1.Buttons.Item(2).ToolTipText = "Mant. Usuarios"
Else
 MDIForm1.SubmenuTit1(8).Caption = "&Tablas"
 MDIForm1.Toolbar1.Buttons.Item(2).ToolTipText = "Mant. TABLAS"
End If
If Wflag = "A" Then
 MDIForm1.Toolbar1.Buttons.Item(13).Enabled = False
End If
ProBar.Value = ProBar.Value + 1
DoEvents
StatusBar1.Enabled = True
DoEvents
MDIForm1.Visible = True
DoEvents
MDIForm1.Show
DoEvents
ProBar.Value = ProBar.Value + 1
Screen.MousePointer = 11
PASSFORM.Hide
WF_REINICIAR = 0
ProBar.Min = 0
ProBar.Visible = False
txtderechos.Visible = False
TXTCIA.Visible = True
TXTCIA.Enabled = True
If TXTCIA.Enabled And TXTCIA.Visible Then TXTCIA.SetFocus
Screen.MousePointer = 0
MDIForm1.SubmenuTit6(6).Enabled = True

Exit Sub
SALE:
   If Err.Number = 360 Then
     Resume Next
   Else
      MsgBox Err.Description, 48, Pub_Titulo
   End If
   Resume Next
Exit Sub
fin_de_pro:
    Screen.MousePointer = 11
    DoEvents
    MDIForm1.WindowState = 1
    MDIForm1.Caption = "Descargando Modulos..."
    EN.Close
    CN.Close
    Screen.MousePointer = 0
    End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If WF_REINICIAR <> 2 Then
    Screen.MousePointer = 11
    MDIForm1.Caption = "Descargando Modulos..."
    DoEvents
    MDIForm1.WindowState = 1
    EN.Close
    CN.Close
    Unload MDIForm1
    Screen.MousePointer = 0
    End
End If
WF_REINICIAR = 0
End Sub
Private Sub menuAlm_Click(Index As Integer)
On Error GoTo SALE
If Trim(menuAlm(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0
Unload FORMGEN
PUB_CODTRA = Val(Left(menuAlm(Index).Caption, 4))
SQ_OPER = 1
LEER_TRA_LLAVE
If Trim(tra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(tra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(tra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
SALE:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
   MsgBox Err.Number
 End If
 Screen.MousePointer = 0
End Sub


Private Sub menuAyudaSistema_Click()
MsgBox "Solution - Gestión Contable " & Chr(13) & "  " & Chr(13) & "Casa de Desarrollo : " & Chr(13) & "UniSoft S.A.C. Trujillo - Perú " & Chr(13) & "Software Empresarial", vbInformation, "UniSoft S.A.C. - Solution - Gestión Contable"
End Sub

Private Sub menuCia_Click(Index As Integer)
Dim iFormCount As Integer
On Error GoTo SIGUE
If Left(MDIForm1.menuCia(Index).Caption, 1) = "*" Then
  Exit Sub
End If
Screen.MousePointer = 11
If Forms.Count - 1 > 0 Then
   For iFormCount = Forms.Count - 1 To 1 Step -1
    If iFormCount <> 1 Then
        Unload Forms(iFormCount)
    End If
   Next iFormCount
End If
For fila = 0 To 20
  MDIForm1.menuCia(fila).Checked = False
Next fila
MDIForm1.menuCia(Index).Checked = True
MDIForm1.StatusBar1.Panels(7).Text = MDIForm1.menuCia(Index).Caption
MDIForm1.TXTCIA.Caption = Trim(MDIForm1.menuCia(Index).Caption)
ACTUALIZA_CIA Left(MDIForm1.menuCia(Index).Caption, 2)
CONTA = 0
PUB_TIPREG = 45
PUB_CODCIA = LK_CODCIA
SQ_OPER = 2
LEER_TAB_LLAVE
Do Until CONTA = 5
    CONTA = CONTA + 1
    If tab_mayor.EOF Then
       tab_precioss(CONTA) = "....."
    Else
       tab_precioss(CONTA) = tab_mayor!tab_nomlargo
       tab_mayor.MoveNext
    End If
Loop
Screen.MousePointer = 0
Exit Sub
SIGUE:
Resume Next
End Sub
Private Sub menuComp_Click(Index As Integer)
On Error GoTo SALE
If Trim(Menucomp(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0
'Unload FORMGEN
PUB_CODTRA = Val(Left(Menucomp(Index).Caption, 4))
SQ_OPER = 1
LEER_TRA_LLAVE
If Trim(tra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(tra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(tra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
SALE:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
   MsgBox Err.Number
 End If
 Resume Next
 Screen.MousePointer = 0
End Sub

Private Sub MenuContab_Click(Index As Integer)
On Error GoTo SALE
If Trim(MenuContab(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0

PUB_CODTRA = Val(Left(MenuContab(Index).Caption, 4))
SQ_OPER = 1
LEER_TRA_LLAVE
If Trim(tra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show 1
ElseIf Trim(tra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(tra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMPC3" Then
 Load FrmImp3
 FrmImp3.Show 1
End If
Screen.MousePointer = 0
Exit Sub
SALE:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
' MsgBox Err.Description
   MsgBox Err.Number
 End If
Screen.MousePointer = 0
End Sub

Private Sub menuctacte_Click()
Screen.MousePointer = 11
Load FrmCaa
FrmCaa.Show
Screen.MousePointer = 0
End Sub

Private Sub menuch_Click()
Screen.MousePointer = 11
Load frm_cheques
Screen.MousePointer = 0
frm_cheques.Show 1

End Sub

Private Sub menudis1_Click(Index As Integer)
On Error GoTo SALE
If Trim(menudis1(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0
Unload FORMGEN
SQ_OPER = 1
PUB_CODTRA = Val(Left(menudis1(Index).Caption, 4))
LEER_TRA_LLAVE
If Trim(tra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(tra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(tra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
SALE:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
   MsgBox Err.Number
 End If
 Screen.MousePointer = 0

End Sub

Private Sub menudis2_Click(Index As Integer)
On Error GoTo SALE
If Trim(menudis2(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0
Unload FORMGEN
SQ_OPER = 1
PUB_CODTRA = Val(Left(menudis2(Index).Caption, 4))
LEER_TRA_LLAVE
If Trim(tra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(tra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(tra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
SALE:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
   MsgBox Err.Number
 End If
 Screen.MousePointer = 0

End Sub

Private Sub menudis3_Click(Index As Integer)
On Error GoTo SALE
If Trim(menudis3(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0
Unload FORMGEN
SQ_OPER = 1
PUB_CODTRA = Val(Left(menudis3(Index).Caption, 4))
LEER_TRA_LLAVE
If Trim(tra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(tra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(tra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
SALE:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
   MsgBox Err.Number
 End If
 Screen.MousePointer = 0

End Sub

Private Sub menuKa_Click()
Screen.MousePointer = 11
Load KARDEX
KARDEX.Show
Screen.MousePointer = 0
End Sub

Private Sub menuMoli_Click(Index As Integer)
If Trim(menuMoli(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0
Unload FORMGEN
PUB_CODTRA = Val(Left(menuMoli(Index).Caption, 4))
SQ_OPER = 1
LEER_TRA_LLAVE
If Trim(tra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(tra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(tra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
End Sub

Private Sub menuRepoAyuda_Click()
Dim wl As Object
Dim var As String

Dim ww As Form
Dim i
MsgBox "En Construcción consulte : UniSoft S.A.C. 044-9669562", 48, Pub_Titulo
 
Exit Sub
On Error GoTo OJO
var = Trim(MDIForm1.ActiveForm.Name)
If UCase(var) = "FORMGEN" Then
  var = Trim(MDIForm1.ActiveForm.Name) & "_" & LK_CODTRA
End If
On Error GoTo 0
If wl Is Nothing Then
    Set wl = CreateObject("word.Application")
End If
var = CONS_ADMIN & "ayudas\" & var
On Error GoTo no_existe
wl.Application.WindowState = wdWindowStateMaximize
'wl.Application.ActiveDocument.Windows(1).WindowState = 2
'ee.Windows(0).WindowState = wdWindowStateMaximize
'wl.Windows(1).WindowState = wdWindowStateMaximize
'wl.Documents
wl.Documents.Open FileName:=var
wl.Application.Visible = True
Set wl = Nothing
On Error GoTo 0
Exit Sub
no_existe:
'Shell "C:\Archivos de programa\Microsoft Office\Office\winword.EXE", "c:\admin\ayudas\FORMGEN_2401.doc""
MsgBox Err.Description
OJO:
If Err.Number <> 91 Then
 MsgBox "Verificar " & Err.Description, 48, Pub_Titulo
ElseIf Err.Number = 91 Then
  var = "MDIForm1"
  On Error GoTo 0
  If wl Is Nothing Then
    Set wl = CreateObject("word.Application")
  End If
  var = CONS_ADMIN & "ayudas\" & var
  On Error GoTo no_existe2
  wl.Application.WindowState = wdWindowStateMaximize
  wl.Documents.Open FileName:=var
  wl.Application.Visible = True
  On Error GoTo 0
End If
no_existe2:
End Sub


Private Sub menuVent_Click(Index As Integer)
On Error GoTo SALE
If Trim(menuVent(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0
'Unload FORMGEN
SQ_OPER = 1
PUB_CODTRA = Val(Left(menuVent(Index).Caption, 4))
LEER_TRA_LLAVE
If Trim(tra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(tra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(tra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show 1
ElseIf Trim(tra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
SALE:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 ElseIf Err.Number = 40009 Then
   MsgBox "No se encontro el Registro del Reporte Reiniciar", 48, Pub_Titulo
 End If

 Screen.MousePointer = 0
End Sub

Private Sub SubmenuTit1_Click(Index As Integer)
Dim h, wBAN As Integer
Dim iFormCount As Integer
On Error GoTo SALE
Select Case Index
    Case 0
        Screen.MousePointer = 11
        Load frmARTI
        frmARTI.Show
        Screen.MousePointer = 0
    Case 1
        Screen.MousePointer = 11
        Load centrocostos
        centrocostos.Show
'        Load frmCLI
'        frmCLI.Show
        Screen.MousePointer = 0
    Case 2
        Screen.MousePointer = 11
        Load FrmVen
        FrmVen.Show
        Screen.MousePointer = 0
    Case 3
        Screen.MousePointer = 11
        Load frmBancos
        Unload FORMGEN
        frmBancos.Show
        Screen.MousePointer = 0
    Case 5
        Screen.MousePointer = 11
        Load frmTablas
        frmTablas.Show
        Screen.MousePointer = 0
    Case 6
        Screen.MousePointer = 11

        If Forms.Count - 1 > 0 Then
           For iFormCount = Forms.Count - 1 To 1 Step -1
            If iFormCount <> 1 Then
                Unload Forms(iFormCount)
            End If
           Next iFormCount
        End If
        Load FrmParGen
        FrmParGen.Show
        Screen.MousePointer = 0
        
    Case 7
 '       Screen.MousePointer = 11
'        Load FrmTiempoReal
'        FrmTiempoReal.Show
'        Screen.MousePointer = 0
        Screen.MousePointer = 0
        Load frmGeneral
        frmGeneral.Show 1
     
    Case 8
        Screen.MousePointer = 11
        Load FrmTabla1
        FrmTabla1.Show
        Screen.MousePointer = 0
    Case 9
        Screen.MousePointer = 11
        Load FrmSub
        FrmSub.Show
        Screen.MousePointer = 0
    Case 10
        Screen.MousePointer = 11
        Load GENCAL
        GENCAL.Show
        Screen.MousePointer = 0
    Case 12
        Screen.MousePointer = 11
        Load FrmClave
        FrmClave.Show
        Screen.MousePointer = 0
    Case 13
       Dim iFrmCount As Integer
       Screen.MousePointer = 11
       For iFrmCount = Forms.Count - 1 To 1 Step -1
         If iFrmCount <> 1 Then
            Unload Forms(iFrmCount)
         End If
       Next iFrmCount
       NOTPermisos
       For fila = 0 To 20
         MDIForm1.menuCia(fila).Checked = False
       Next fila
       PUB_FLAG = 1
       Screen.MousePointer = 11
       MDIForm_Load
       Screen.MousePointer = 0
    Case 14
        'REINICIAR SISTEMA
        WF_REINICIAR = 2 'PARA VERIFICAR EN MDIFORM_UNLOAD
        Screen.MousePointer = 11
        MDIForm1.Caption = "Descargando Modulos..."
        CN.Cancel
        CN.Close
        MDIForm1.WindowState = 1
        MDIForm1.Hide
        Unload MDIForm1
        PUB_FLAG = 0
        Unload Splash
        Screen.MousePointer = 0
        Load Splash
    Case 15
        'SALIR
        Screen.MousePointer = 11
        DoEvents
        MDIForm1.WindowState = 1
        MDIForm1.Caption = "Descargando Modulos..."
        EN.Close
        CN.Close
        Screen.MousePointer = 0
        End
    Case Else
        MsgBox "INDICE FUERA DE RANGO ... Favor de llamar a Computo ", 48, Pub_Titulo
End Select
Exit Sub
SALE:
'MsgBox Err.Description
If Err.Number = 424 Then
 Screen.MousePointer = 0
 MsgBox "No tiene Acceso ", 48, Pub_Titulo
End If

End Sub

Private Sub SubmenuTit2_Click(Index As Integer)
On Error GoTo SALE
Select Case Index
    Case 0
        Screen.MousePointer = 11
        Load FORMGEN
        FORMGEN.Show
        Screen.MousePointer = 0
    Case 1
       Screen.MousePointer = 11
       Load Costos
       Costos.Show
       Screen.MousePointer = 0
    Case 3
       Screen.MousePointer = 11
       Load CIERRE_DIA
       Screen.MousePointer = 0
       CIERRE_DIA.Show 1
    Case Else
        MsgBox "INDICE FUERA DE RANGO ...Favor de llamar a Computo ..!1", 48, Pub_Titulo
End Select
Exit Sub
SALE:
If Err.Number = 424 Then
 Screen.MousePointer = 0
 MsgBox "No tiene Acceso ", 48, Pub_Titulo
End If

End Sub


Private Sub SubmenuTit6_Click(Index As Integer)
On Error GoTo SALE
Select Case Index
    Case 0
     Exit Sub
    Case 1
        Screen.MousePointer = 11
        Load frmdocu
        Screen.MousePointer = 0
        frmdocu.Show
    Case 2
        Screen.MousePointer = 11
        Load FrmPreUnidad
        FrmPreUnidad.Show
        Screen.MousePointer = 0
    Case 4
        Screen.MousePointer = 11
        Load FrmTiempoReal
        Screen.MousePointer = 0
        FrmTiempoReal.Show 1

    Case 8
       Load frmPesos
       frmPesos.Show 1
       If LKCHEK = False Then
         GoTo sale2
        End If
        Dim ArchOrigen, ArchDestino As String
        Dim momento As String
        Screen.MousePointer = 0
        pub_mensaje = "Inserte el Disco de la Portatil en la Unidad A: , Luego Aceptar.."
        Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
        If Pub_Respuesta = vbYes Then
          Screen.MousePointer = 0
          Exit Sub
        End If
        Screen.MousePointer = 11
disco:
        On Error GoTo sale2
        
        ArchOrigen = CONS_ADMIN & "PEDIDOS\VER.TXT"
        ArchDestino = "a:\VER"
        FileCopy ArchOrigen, ArchDestino
        
        Call Shell(CONS_ADMIN & "PEDIDOS\COPIAR.PIF", vbNormalFocus)
        momento = Timer + 3#
        Do While momento > Timer
         DoEvents
        Loop
        Screen.MousePointer = 0
        'msgbox "Proceso Terminado..", 48, pub_titulo
        Exit Sub
sale2:
        Screen.MousePointer = 0
    Case 6
       Screen.MousePointer = 11
       Load FORM_PED
       FORM_PED.Show
       Screen.MousePointer = 0
    Case Else
        MsgBox "INDICE FUERA DE RANGO ... Favor de llamar a Computo ", 48, Pub_Titulo
End Select
Exit Sub
SALE:
Screen.MousePointer = 0
If Err.Number = 53 Then
    MsgBox "NO esta Instalado  ...", 48, Pub_Titulo
    Exit Sub
End If
If Err.Number = 424 Then
 Screen.MousePointer = 0
 MsgBox "No tiene Acceso ", 48, Pub_Titulo
End If

End Sub

Private Sub SubmenuTit5_Click(Index As Integer)
On Error GoTo SALE
Select Case Index
    Case 0
       Call Shell("C:\MSSQL\BINN\SQLEW.EXE", 3) ' Ejecuta SQL
    Case 1
        Load frmeditor
        frmeditor.Show 1
    Case 2
       Screen.MousePointer = 11
       Unload FORMGEN
       Load frmConfigRepo
       Screen.MousePointer = 0
       frmConfigRepo.Show 1
    Case 3
        If LK_CODUSU <> "ADMIN" Then
          Load FrmGrupos
          FrmGrupos.Show 1
        Else
          pub_mensaje = "ADMIN , desea el Formulario de Imprimer BDATOS "
          Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
          If Pub_Respuesta = vbYes Then
            Load FrmImprime
            FrmImprime.Show 1
          Else
            Load FrmGrupos
            FrmGrupos.Show 1
          End If
       End If
    Case 4
        Load frmDiseno
        frmDiseno.Show 1
    Case Else
        MsgBox "INDICE FUERA DE RANGO ...Favor de llamar a Computo ..!1", 48, Pub_Titulo
End Select

Exit Sub
SALE:
If Err.Number = 53 Then
 Screen.MousePointer = 0
 MsgBox "Si Elijio Backup, Este No Procede en Terminales. Solo en su Servidor , Consulte a su Proveedor para mas Detalles", 48, Pub_Titulo
 Exit Sub
End If
If Err.Number = 424 Then
 Screen.MousePointer = 0
 MsgBox "No tiene Acceso ", 48, Pub_Titulo
End If
End Sub
Private Sub SubmenuTit3_Click(Index As Integer)
On Error GoTo SALE
Select Case Index
    Case 0
        Screen.MousePointer = 11
        Load FrmCoPar
        Screen.MousePointer = 0
        FrmCoPar.Show 1
    Case 1
        Screen.MousePointer = 11
        Load FORM_CONT
        FORM_CONT.Show
        Screen.MousePointer = 0
    Case 2
        Screen.MousePointer = 11
        Load FrmPlan
        Screen.MousePointer = 0
        FrmPlan.Show
    Case 4
        Screen.MousePointer = 11
        Load FORM_CONTA2
        FORM_CONTA2.Show
        Screen.MousePointer = 0
    Case 5
        Screen.MousePointer = 11
        Load FORM_CONTA
        FORM_CONTA.Show
        Screen.MousePointer = 0
    Case 6
        Screen.MousePointer = 11
        Load FrmProce
        Screen.MousePointer = 0
        FrmProce.Show 1
    Case 7
        'Screen.MousePointer = 11
        'Load FrmProce
        'Screen.MousePointer = 0
        'FrmProce.Show 1
        MsgBox "En Construccion..."
    Case 9
        Screen.MousePointer = 11
        Load FrmCaa
        Screen.MousePointer = 0
        FrmCaa.Show
       
    Case 11
        Screen.MousePointer = 11
        Load frmExportSC ' GenContab
        Screen.MousePointer = 0
       
       frmExportSC.Show 1
    Case 12
        Screen.MousePointer = 11
        Load centrocostos
        Screen.MousePointer = 0
        centrocostos.Show
    Case 13
        Screen.MousePointer = 11
        Load frm_mayoriz
        Screen.MousePointer = 0
        frm_mayoriz.Show 1
    Case 14
        Screen.MousePointer = 11
        Load frm_mes_cierre
        Screen.MousePointer = 0
        frm_mes_cierre.Show 1
    Case 16
        Screen.MousePointer = 11
        Load frmimport
        Screen.MousePointer = 0
        frmimport.Show 1
    Case Else
        MsgBox "INDICE FUERA DE RANGO ...Favor de llamar a Computo ..!1", 48, Pub_Titulo
End Select
Exit Sub
SALE:
If Err.Number = 424 Then
 Screen.MousePointer = 0
 MsgBox "No tiene Acceso ", 48, Pub_Titulo
End If


End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
On Error GoTo sal
Dim iFormCount As Integer
Select Case Button.Key
  Case "FORMGEN"
      Screen.MousePointer = 11
      Load FORMGEN
      FORMGEN.Show
      Screen.MousePointer = 0
  Case "Tablas"
      Screen.MousePointer = 11
      Load FrmTabla1
      FrmTabla1.Show
      Screen.MousePointer = 0
  Case "KARDEX"
      Screen.MousePointer = 11
      Load KARDEX
      KARDEX.Show
      Screen.MousePointer = 0
  Case "PRECIOS"
      Screen.MousePointer = 11
      Load FrmPreUnidad
      FrmPreUnidad.Show
      Screen.MousePointer = 0
  Case "ARTI"
      Screen.MousePointer = 11
      Load frmARTI
      frmARTI.Show
      Screen.MousePointer = 0
  Case "CLIENTES"
      Screen.MousePointer = 11
      Load frmCLI
      frmCLI.Show
      Screen.MousePointer = 0
  Case "CUENTAS"
      Load FrmCaa
      FrmCaa.Show
  Case "CHEQUES"
      Load frm_cheques
      frm_cheques.Show 1
  Case "Password"
      Load FrmClave
      FrmClave.Show 1
  Case "BANCOS"
      Load frmBancos
      frmBancos.Show
  Case "Salir"
        Screen.MousePointer = 11
        DoEvents
        MDIForm1.WindowState = 1
        MDIForm1.Caption = "Descargando Modulos..."
        EN.Close
        CN.Close
        Screen.MousePointer = 0
        End
  Case "ALTERNO"
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      Toolbar1.Buttons.Item(13).ToolTipText = "Modo Original"
      LK_FLAG_ALTERNO = " "
      Toolbar1.Buttons.Item(13).Image = 48
      'DoEvents
    Else
      Toolbar1.Buttons.Item(13).ToolTipText = "Modo Alterno"
      LK_FLAG_ALTERNO = "A"
      Toolbar1.Buttons.Item(13).Image = 56
     ' DoEvents
    End If
    'SQ_OPER = 1
    'PUB_CODCIA = LK_CODCIA
    'LEER_PAR_LLAVE
    'If par_llave.EOF Then
    '    MsgBox "NO Existe Compañia Avisar !!!! Posible Error .. Intente Nuevamente ....", 48, Pub_Titulo
    '    End
    '    Exit Sub
    'End If
    'par_llave.Edit
    'par_llave!PAR_FLAG_ALTERNO = LK_FLAG_ALTERNO
    'par_llave.Update
   If Forms.Count - 1 > 0 Then
     For iFormCount = Forms.Count - 1 To 1 Step -1
        If iFormCount <> 1 Then
           If UCase(Forms(iFormCount).Name) = "FRMARTI" Then
                'MsgBox UCase(Forms(iFormCount).Name)
                frmARTI.PROCESO_CANCELAR
                frmARTI.PROCESO_ARTI
           ElseIf UCase(Forms(iFormCount).Name) = "KARDEX" Then
                KARDEX.ListView1.Visible = False
                KARDEX.i_codart2.Text = ""
           ElseIf UCase(Forms(iFormCount).Name) = "FRMPREUNIDAD" Then
                FrmPreUnidad.PRO_GOTFOCUS
                FrmPreUnidad.ListView1.Visible = False
                FrmPreUnidad.txt_key.Text = ""
           ElseIf UCase(Forms(iFormCount).Name) = "FRMREPO" Then
                FrmRepo.ListView1.Visible = False
                FrmRepo.txt_key.Text = ""
           ElseIf UCase(Forms(iFormCount).Name) = "FORMGEN" Then
                FORMGEN.LV_ART.Visible = False
                FORMGEN.TEXTOVAR.Text = ""
           End If
        End If
       Next iFormCount
   End If

  Case "Excel"
     Call Shell("c:\Archivos de programa\Microsoft Office\Office\EXCEL.EXE", 3)    ' Ejecuta Excel
     'Call Shell("c:\Archivos de programa\Microsoft Office\Office\EXCEL.EXE", 3)    ' Ejecuta Excel
     
End Select
Exit Sub

sal:
  
  Select Case Err.Number
  Case 53
    MsgBox "No se encontro el Archivo ....!!!!..", 48, Pub_Titulo
  Case 424
    MsgBox "No existe el Formulario ....!!!!..r", 48, Pub_Titulo
  Case Else
  
    MsgBox Err.Number & "  " & Err.Description & "    !!!!... ", 48, Pub_Titulo
  End Select
Screen.MousePointer = 0
  Exit Sub


End Sub

Private Sub TXTCIA_Click()
  MDIForm1.PopupMenu Menucompania
End Sub

