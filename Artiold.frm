VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmARTI 
   Caption         =   "Maestro de Articulos"
   ClientHeight    =   6750
   ClientLeft      =   690
   ClientTop       =   1185
   ClientWidth     =   11325
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cdlfoto 
      Left            =   2760
      Top             =   8280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdpath 
      Caption         =   "..."
      Height          =   255
      Left            =   13440
      TabIndex        =   139
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtpath 
      Height          =   285
      Left            =   11880
      TabIndex        =   138
      Top             =   6600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox picfoto 
      Height          =   2055
      Left            =   11880
      ScaleHeight     =   1995
      ScaleWidth      =   1875
      TabIndex        =   137
      Top             =   4320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   540
      Left            =   9645
      TabIndex        =   136
      Top             =   7770
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame frmbusqueda 
      BackColor       =   &H00C0C0C0&
      Height          =   4110
      Left            =   240
      TabIndex        =   125
      Tag             =   "9898"
      Top             =   4200
      Visible         =   0   'False
      Width           =   11430
      Begin VB.ComboBox artfamilia 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   150
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   129
         Top             =   450
         Width           =   2715
      End
      Begin VB.ComboBox artsubfam 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   2925
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   128
         Top             =   465
         Width           =   2715
      End
      Begin VB.ComboBox artgrupo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   5685
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   127
         Top             =   465
         Width           =   2715
      End
      Begin VB.ComboBox artlinea 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   8445
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   126
         Top             =   480
         Width           =   2715
      End
      Begin MSFlexGridLib.MSFlexGrid grdarticulos 
         Height          =   3135
         Left            =   60
         TabIndex        =   130
         Top             =   915
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5530
         _Version        =   393216
         BackColorBkg    =   16777215
         SelectionMode   =   1
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "División:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F8DED7&
         Height          =   195
         Index           =   15
         Left            =   180
         TabIndex        =   134
         Top             =   195
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Familia:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F8DED7&
         Height          =   195
         Index           =   16
         Left            =   2940
         TabIndex        =   133
         Top             =   210
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Linea:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F8DED7&
         Height          =   195
         Index           =   17
         Left            =   5685
         TabIndex        =   132
         Top             =   225
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F8DED7&
         Height          =   195
         Index           =   18
         Left            =   8490
         TabIndex        =   131
         Top             =   225
         Width           =   495
      End
      Begin VB.Label lblart 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   732
         Index           =   7
         Left            =   0
         TabIndex        =   135
         Top             =   0
         Width           =   11316
      End
   End
   Begin VB.CommandButton cmd_AddItem 
      Caption         =   "Insertar Items"
      Height          =   270
      Left            =   10440
      TabIndex        =   106
      Top             =   3435
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.CommandButton MANOS 
      Height          =   405
      Index           =   0
      Left            =   9360
      Picture         =   "Arti.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   735
      Width           =   420
   End
   Begin VB.CommandButton MANOS 
      Height          =   405
      Index           =   1
      Left            =   9825
      Picture         =   "Arti.frx":0702
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   735
      Width           =   420
   End
   Begin VB.Frame F14 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2652
      Left            =   480
      TabIndex        =   59
      Top             =   5040
      Visible         =   0   'False
      Width           =   10812
      Begin VB.OptionButton Op 
         Caption         =   "Actualizar items seleccionados"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   64
         Top             =   1800
         Width           =   2535
      End
      Begin VB.OptionButton Op 
         Caption         =   "Ignorar la Lista "
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   63
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton cmdconfirma 
         Caption         =   "Con&firmar Grabación"
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
         Left            =   4320
         TabIndex        =   62
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CommandButton CmdEscapa 
         Caption         =   "E&scapar"
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
         Left            =   6720
         TabIndex        =   61
         Top             =   1800
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid ListExiste 
         Height          =   1536
         Left            =   48
         TabIndex        =   60
         Top             =   192
         Width           =   7812
         _ExtentX        =   13785
         _ExtentY        =   2699
         _Version        =   393216
         Cols            =   5
         BackColorFixed  =   9128212
         ForeColorFixed  =   16777215
         BackColorBkg    =   16777215
         GridColorFixed  =   12632256
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Utilize la barra espaciadora para marcar o desmarcar un item."
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
         Left            =   0
         TabIndex        =   144
         Top             =   0
         Width           =   7575
      End
   End
   Begin VB.TextBox tcospro 
      DataField       =   "ART_COSPRO"
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
      Height          =   300
      Left            =   5520
      MaxLength       =   11
      TabIndex        =   79
      Top             =   810
      Visible         =   0   'False
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   0
      TabIndex        =   32
      Top             =   840
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   5106
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Definición de Estructura"
      TabPicture(0)   =   "Arti.frx":1044
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fvarios"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "pgb_Progress"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Porcentajes"
      TabPicture(1)   =   "Arti.frx":1060
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CERO"
      Tab(1).Control(1)=   "Fop"
      Tab(1).Control(2)=   "Fcomi"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Almacen Defectuosos"
      TabPicture(2)   =   "Arti.frx":107C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frarelacion"
      Tab(2).Control(1)=   "frmpro"
      Tab(2).ControlCount=   2
      Begin ComctlLib.ProgressBar pgb_Progress 
         Height          =   210
         Left            =   6960
         TabIndex        =   118
         Top             =   -15
         Visible         =   0   'False
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   370
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.Frame frarelacion 
         Caption         =   "Articulo Relacionado con Almacen Defectuoso : "
         Height          =   2415
         Left            =   -74880
         TabIndex        =   93
         Top             =   360
         Width           =   9975
         Begin VB.TextBox txtcodigo2 
            Height          =   285
            Left            =   7320
            TabIndex        =   97
            Top             =   1200
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Adicionar"
            Height          =   750
            Left            =   7260
            Picture         =   "Arti.frx":1098
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   1575
            Width           =   1140
         End
         Begin VB.CommandButton cmdquitar 
            Caption         =   "&Quitar Relación"
            Height          =   765
            Left            =   8520
            Picture         =   "Arti.frx":14DA
            Style           =   1  'Graphical
            TabIndex        =   95
            Top             =   1560
            Width           =   1335
         End
         Begin VB.ComboBox cmbcal 
            Height          =   315
            Left            =   7320
            Style           =   2  'Dropdown List
            TabIndex        =   94
            Top             =   600
            Width           =   2295
         End
         Begin MSFlexGridLib.MSFlexGrid gridrel 
            Height          =   1935
            Left            =   120
            TabIndex        =   96
            Top             =   240
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   3413
            _Version        =   393216
            BackColorFixed  =   9128212
            ForeColorFixed  =   16777215
            BackColorBkg    =   16777215
            GridColorFixed  =   12632256
            GridLinesFixed  =   1
            Appearance      =   0
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Codigo de Relación"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   7200
            TabIndex        =   99
            Top             =   960
            Visible         =   0   'False
            Width           =   1485
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblcal 
            Caption         =   "Calidad Relacionada para Agregar"
            Height          =   375
            Left            =   7200
            TabIndex        =   98
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame frmpro 
         Caption         =   "Relación de Procesos"
         Height          =   375
         Left            =   -74640
         TabIndex        =   76
         Top             =   3360
         Width           =   8415
         Begin VB.Data dataO 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   375
            Left            =   360
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   1560
            Width           =   1140
         End
         Begin VB.CommandButton cmdp 
            Caption         =   "Activar Relación"
            Height          =   735
            Left            =   240
            TabIndex        =   78
            Top             =   600
            Width           =   1455
         End
         Begin MSFlexGridLib.MSFlexGrid gridp 
            Height          =   2055
            Left            =   1800
            TabIndex        =   77
            Top             =   240
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   3625
            _Version        =   393216
            Cols            =   4
         End
      End
      Begin VB.CommandButton CERO 
         Caption         =   "Producto 0 (Para ADMIN)"
         Height          =   285
         Left            =   -67500
         TabIndex        =   67
         Top             =   2385
         Width           =   2415
      End
      Begin VB.Frame Fop 
         Caption         =   "Opciones"
         ForeColor       =   &H00C00000&
         Height          =   2415
         Left            =   -71760
         TabIndex        =   58
         Top             =   360
         Width           =   6855
         Begin VB.ComboBox decimales 
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
            Left            =   5160
            TabIndex        =   29
            ToolTipText     =   "Decimales para la Cantidad Formulada"
            Top             =   720
            Width           =   1455
         End
         Begin VB.ComboBox art_codpro 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   240
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1440
            Width           =   2655
         End
         Begin VB.TextBox txtfechault 
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
            Left            =   5160
            TabIndex        =   28
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtMax 
            Height          =   285
            Left            =   1800
            MaxLength       =   13
            TabIndex        =   26
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtMin 
            Height          =   285
            Left            =   1800
            MaxLength       =   13
            TabIndex        =   25
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Decimales:"
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
            Index           =   19
            Left            =   3120
            TabIndex        =   74
            Top             =   720
            Width           =   765
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ultima Compra de:"
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
            Left            =   240
            TabIndex        =   69
            Top             =   1200
            Width           =   1320
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Ultima Compra"
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
            Index           =   0
            Left            =   3120
            TabIndex        =   68
            Top             =   360
            Width           =   1515
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Stock Minimo :"
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
            Index           =   1
            Left            =   240
            TabIndex        =   66
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Stock Maximo :"
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
            Index           =   2
            Left            =   240
            TabIndex        =   65
            Top             =   720
            Width           =   1080
         End
      End
      Begin VB.Frame Fvarios 
         Height          =   2445
         Left            =   75
         TabIndex        =   47
         Top             =   360
         Width           =   10050
         Begin VB.TextBox txtmarca 
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
            Left            =   3720
            MaxLength       =   30
            TabIndex        =   141
            Top             =   1477
            Width           =   3375
         End
         Begin VB.TextBox txtcolor 
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
            Left            =   3720
            MaxLength       =   60
            TabIndex        =   140
            Top             =   2040
            Width           =   3375
         End
         Begin VB.OptionButton cheservi 
            BackColor       =   &H00808080&
            Caption         =   "Mercaderia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Index           =   0
            Left            =   7305
            TabIndex        =   116
            Top             =   1845
            Width           =   1335
         End
         Begin VB.OptionButton cheservi 
            BackColor       =   &H00808080&
            Caption         =   "Paquete de Mercaderia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Index           =   1
            Left            =   7440
            TabIndex        =   115
            Top             =   2115
            Width           =   2310
         End
         Begin VB.OptionButton cheservi 
            BackColor       =   &H00808080&
            Caption         =   "Servicio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Index           =   2
            Left            =   8595
            TabIndex        =   114
            Top             =   1845
            Width           =   1185
         End
         Begin VB.ComboBox DS 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   315
            Left            =   8625
            Style           =   2  'Dropdown List
            TabIndex        =   111
            Top             =   1395
            Width           =   840
         End
         Begin VB.CheckBox art_situacion 
            BackColor       =   &H00808080&
            Caption         =   "Desactivado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   7395
            TabIndex        =   110
            Top             =   210
            Width           =   1935
         End
         Begin VB.CheckBox checambio 
            BackColor       =   &H00808080&
            Caption         =   "Para Cambio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   7395
            TabIndex        =   109
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox txtcospro 
            DataField       =   "ART_COSPRO"
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
            Height          =   300
            Left            =   8220
            MaxLength       =   11
            TabIndex        =   108
            Top             =   1050
            Width           =   1215
         End
         Begin VB.CheckBox exigv 
            BackColor       =   &H00808080&
            Caption         =   "Exoneración IGV"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   195
            Left            =   7395
            TabIndex        =   107
            Top             =   750
            Width           =   2250
         End
         Begin VB.ComboBox art_plancha 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   315
            Left            =   4920
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   720
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.ComboBox art_marca 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   315
            Left            =   3720
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   960
            Width           =   3375
         End
         Begin VB.ComboBox art_linea 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   315
            Left            =   3720
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   315
            Width           =   3375
         End
         Begin VB.ComboBox art_numero 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   315
            Left            =   80
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2060
            Width           =   3375
         End
         Begin VB.ComboBox art_grupo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   315
            Left            =   80
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1477
            Width           =   3375
         End
         Begin VB.ComboBox art_subfam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   315
            Left            =   80
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   896
            Width           =   3375
         End
         Begin VB.ComboBox art_familia 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   315
            Left            =   80
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   315
            Width           =   3375
         End
         Begin VB.ComboBox CmbCalidad 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   315
            Left            =   4320
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   0
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.Label Label9 
            Caption         =   "Marca:"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   3720
            TabIndex        =   143
            Top             =   1280
            Width           =   1965
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda Articulo:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Index           =   18
            Left            =   7395
            TabIndex        =   113
            Top             =   1470
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(%) Igv :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Index           =   10
            Left            =   7395
            TabIndex        =   112
            Top             =   1110
            Width           =   675
         End
         Begin VB.Label Lbl3 
            AutoSize        =   -1  'True
            Caption         =   "Color"
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
            Left            =   3720
            TabIndex        =   101
            Top             =   1860
            Width           =   375
         End
         Begin VB.Label lblart 
            Caption         =   "Lote:"
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
            Index           =   6
            Left            =   3720
            TabIndex        =   80
            Top             =   795
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label lblart 
            Caption         =   "Clase:"
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
            Index           =   5
            Left            =   3720
            TabIndex        =   75
            Top             =   700
            Width           =   1965
         End
         Begin VB.Label lblart 
            Caption         =   "Lote:"
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
            Index           =   4
            Left            =   3720
            TabIndex        =   73
            Top             =   120
            Width           =   3000
         End
         Begin VB.Label lblart 
            Caption         =   "Sub Linea:"
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
            Index           =   3
            Left            =   120
            TabIndex        =   72
            Top             =   1860
            Width           =   1965
         End
         Begin VB.Label lblart 
            Caption         =   "Linea:"
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
            Index           =   2
            Left            =   120
            TabIndex        =   71
            Top             =   1280
            Width           =   1965
         End
         Begin VB.Label lblart 
            Caption         =   "Familia:"
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
            Index           =   1
            Left            =   120
            TabIndex        =   70
            Top             =   700
            Width           =   3330
         End
         Begin VB.Label lblart 
            AutoSize        =   -1  'True
            Caption         =   "División:"
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
            Index           =   0
            Left            =   120
            TabIndex        =   48
            Top             =   120
            Width           =   600
         End
         Begin VB.Label Label8 
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Height          =   2250
            Left            =   7230
            TabIndex        =   123
            Top             =   150
            Width           =   2685
         End
      End
      Begin VB.Frame Fcomi 
         Caption         =   "Porc. de Comisiones por Articulo"
         ForeColor       =   &H00C00000&
         Height          =   2415
         Left            =   -74880
         TabIndex        =   52
         Top             =   360
         Visible         =   0   'False
         Width           =   3015
         Begin VB.TextBox txtpor6 
            Height          =   285
            Left            =   1560
            MaxLength       =   6
            TabIndex        =   120
            Top             =   2040
            Width           =   1095
         End
         Begin VB.TextBox txtpor5 
            Height          =   285
            Left            =   1560
            MaxLength       =   6
            TabIndex        =   24
            Top             =   1695
            Width           =   1095
         End
         Begin VB.TextBox txtpor4 
            Height          =   285
            Left            =   1560
            MaxLength       =   6
            TabIndex        =   23
            Top             =   1350
            Width           =   1095
         End
         Begin VB.TextBox txtpor3 
            Height          =   285
            Left            =   1560
            MaxLength       =   6
            TabIndex        =   22
            Top             =   990
            Width           =   1095
         End
         Begin VB.TextBox txtpor2 
            Height          =   285
            Left            =   1560
            MaxLength       =   6
            TabIndex        =   21
            Top             =   645
            Width           =   1095
         End
         Begin VB.TextBox txtpor1 
            Height          =   285
            Left            =   1560
            MaxLength       =   6
            TabIndex        =   20
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label lblpor 
            AutoSize        =   -1  'True
            Caption         =   "% p'  Precios 6 :"
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
            Index           =   5
            Left            =   105
            TabIndex        =   119
            Top             =   2085
            Width           =   1170
         End
         Begin VB.Label lblpor 
            AutoSize        =   -1  'True
            Caption         =   "% p'  Precios 5 :"
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
            Index           =   4
            Left            =   120
            TabIndex        =   57
            Top             =   1728
            Width           =   1170
         End
         Begin VB.Label lblpor 
            AutoSize        =   -1  'True
            Caption         =   "% Dscto.25 - +"
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
            Index           =   3
            Left            =   120
            TabIndex        =   56
            Top             =   1365
            Width           =   1125
         End
         Begin VB.Label lblpor 
            AutoSize        =   -1  'True
            Caption         =   "% Dscto.0-24.99"
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
            Index           =   2
            Left            =   120
            TabIndex        =   55
            Top             =   1020
            Width           =   1245
         End
         Begin VB.Label lblpor 
            AutoSize        =   -1  'True
            Caption         =   "% Credito:"
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
            Index           =   1
            Left            =   120
            TabIndex        =   54
            Top             =   660
            Width           =   795
         End
         Begin VB.Label lblpor 
            AutoSize        =   -1  'True
            Caption         =   "% Contado:"
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
            Index           =   0
            Left            =   135
            TabIndex        =   53
            Top             =   300
            Width           =   885
         End
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Correla"
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
         Left            =   -68640
         TabIndex        =   46
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Nº. Dir."
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
         Left            =   -71040
         TabIndex        =   45
         Top             =   480
         Width           =   645
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Direc. Trabajo :"
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
         Left            =   -74760
         TabIndex        =   44
         Top             =   480
         Width           =   1350
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "SubZona Trab."
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
         Left            =   -70920
         TabIndex        =   43
         Top             =   1080
         Width           =   1290
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Zona : Trab."
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
         Left            =   -74760
         TabIndex        =   42
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Prendas :"
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
         Left            =   -71880
         TabIndex        =   41
         Top             =   3480
         Width           =   825
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Prop. (1) :"
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
         Left            =   -74760
         TabIndex        =   40
         Top             =   1680
         Width           =   870
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Prop. (2) :"
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
         Left            =   -74760
         TabIndex        =   39
         Top             =   2280
         Width           =   870
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Rg.Pub.(1)"
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
         Left            =   -71760
         TabIndex        =   38
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Rg.Pub.(2)"
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
         Left            =   -71880
         TabIndex        =   37
         Top             =   2280
         Width           =   930
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Autovaluo :"
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
         Left            =   -74760
         TabIndex        =   36
         Top             =   3480
         Width           =   990
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Autos (1) :"
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
         Left            =   -74760
         TabIndex        =   35
         Top             =   2880
         Width           =   900
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Autos (2) :"
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
         Left            =   -71760
         TabIndex        =   34
         Top             =   2880
         Width           =   900
      End
      Begin VB.Label Label2 
         Caption         =   "    Relación      Cia  -  Cuenta"
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
         Index           =   2
         Left            =   -67080
         TabIndex        =   33
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame Fcomun 
      Height          =   3390
      Left            =   0
      TabIndex        =   81
      Top             =   3765
      Visible         =   0   'False
      Width           =   11835
      Begin VB.TextBox txtlitro 
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5610
         TabIndex        =   11
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox txtvar 
         Height          =   285
         Left            =   6240
         MaxLength       =   9
         TabIndex        =   82
         Top             =   1095
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmddolares 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5610
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   510
         Width           =   5595
      End
      Begin VB.TextBox txtpeso 
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2160
         TabIndex        =   10
         Top             =   180
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid grid_unid 
         Height          =   2220
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "[INSERT] Agrega, [DEL] Quitar"
         Top             =   1095
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   3916
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         BackColorFixed  =   9128212
         ForeColorFixed  =   16777215
         BackColorBkg    =   16777215
         GridColorFixed  =   12632256
         Enabled         =   0   'False
         FocusRect       =   2
         HighLight       =   2
         GridLinesFixed  =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRECIO. 6"
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
         Height          =   240
         Index           =   5
         Left            =   9795
         TabIndex        =   117
         Tag             =   "6"
         Top             =   855
         Width           =   1395
      End
      Begin VB.Label Label7 
         Caption         =   "Litros Unid. Act :"
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
         Height          =   255
         Left            =   4185
         TabIndex        =   103
         Top             =   195
         Width           =   1395
      End
      Begin VB.Label lblcospro 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   10110
         TabIndex        =   12
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label LOTRO 
         Caption         =   "C. Prom.  Unidad Act. : "
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
         Height          =   255
         Left            =   7950
         TabIndex        =   92
         Top             =   180
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Unidad Activa :"
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
         Index           =   0
         Left            =   120
         TabIndex        =   91
         Top             =   540
         Width           =   1260
      End
      Begin VB.Label LBLCOSTO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "COSTO"
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
         Height          =   240
         Left            =   2025
         TabIndex        =   90
         Top             =   855
         Width           =   780
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRECIO. 5"
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
         Height          =   240
         Index           =   4
         Left            =   8415
         TabIndex        =   89
         Tag             =   "5"
         Top             =   855
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRECIO. 4"
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
         Height          =   240
         Index           =   3
         Left            =   7005
         TabIndex        =   88
         Tag             =   "4"
         Top             =   855
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRECIO. 3"
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
         Height          =   240
         Index           =   2
         Left            =   5625
         TabIndex        =   87
         Tag             =   "3"
         Top             =   855
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRECIO. 2"
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
         Height          =   240
         Index           =   1
         Left            =   4215
         TabIndex        =   86
         Tag             =   "2"
         Top             =   855
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRECIO. 1"
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
         Height          =   240
         Index           =   0
         Left            =   2835
         TabIndex        =   85
         Tag             =   "1"
         Top             =   855
         Width           =   1400
      End
      Begin VB.Label LBLUNIDAD 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "UNIDAD"
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
         Height          =   285
         Left            =   120
         TabIndex        =   84
         Top             =   780
         Width           =   1290
      End
      Begin VB.Label Label5 
         Caption         =   "Peso(Kg) Unidad Act.:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   180
         Width           =   1935
      End
   End
   Begin VB.Frame Fdatos 
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
      Height          =   700
      Left            =   30
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   10200
      Begin VB.TextBox txtnombre 
         DataField       =   "ART_NOMBRE"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2865
         MaxLength       =   60
         TabIndex        =   121
         Top             =   300
         Width           =   6975
      End
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
         MaxLength       =   20
         TabIndex        =   0
         Top             =   315
         Width           =   1095
      End
      Begin VB.TextBox txt_alterno 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   1
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label lblnomarti 
         AutoSize        =   -1  'True
         Caption         =   "Descripción del Articulo"
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
         Left            =   2865
         TabIndex        =   122
         Top             =   120
         Width           =   1980
      End
      Begin VB.Label lblalterno 
         AutoSize        =   -1  'True
         Caption         =   "Sub-Codigo"
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
         Left            =   1440
         TabIndex        =   51
         Top             =   120
         Width           =   960
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Interno:"
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
         Index           =   20
         Left            =   120
         TabIndex        =   50
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Limpiar"
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
      Height          =   435
      Left            =   10455
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1935
      Width           =   1300
   End
   Begin VB.Timer Parpadea 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   7320
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Ce&rrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10455
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2490
      Width           =   1300
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10455
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   270
      Width           =   1300
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10455
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1380
      Width           =   1300
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Adicionar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10455
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   825
      Width           =   1300
   End
   Begin VB.Label lblart 
      Caption         =   "Linea:"
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
      Index           =   8
      Left            =   0
      TabIndex        =   142
      Top             =   0
      Width           =   1965
   End
   Begin VB.Label lblprogress 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   10455
      TabIndex        =   124
      Top             =   3075
      Width           =   1290
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "OSBusiness"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   102
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4914&
      BorderStyle     =   1  'Fixed Single
      Height          =   3750
      Index           =   7
      Left            =   10320
      TabIndex        =   100
      Top             =   15
      Width           =   1575
   End
   Begin VB.Label LblMensaje 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4320
      TabIndex        =   31
      Top             =   6645
      Width           =   3285
   End
End
Attribute VB_Name = "frmARTI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pres As String
Dim updatePrecios As Boolean
'**************************
Dim CU As Integer
Dim loc_flag_bloq As String * 1
Dim PSART_LOC As rdoQuery
Dim PSART_LOC2 As rdoQuery
Dim artloc_llave As rdoResultset
Dim artloc_mayor As rdoResultset
Dim PSART_KEY As rdoQuery
Dim artloc_key As rdoResultset
Dim LOC_OPER As Integer
Dim loc_tipo As String * 1
Dim LOC_TIPREG As Integer
Dim temporal As String
Dim loc_key As Integer
Dim Flag_Consis As String * 1
Dim Flag_F2 As String * 1
Dim Flag_Bloq As String * 1
Dim Flag_Inicial As String * 1
Dim Flag_Change  As String * 1
Dim loc_fila  As Integer
Dim loc_colum  As Integer
Dim loc_unid As String * 1
Dim VAR_ACTIVAR As Integer
Dim LOC_ORIGINAL As Currency
Dim LOC_ALTERNO As String
Dim LOC_NOMBRE As String
Dim LOC_CALIDAD As Integer
Dim VAR_NEWCAL As Integer
Dim LOC_CODART2 As Currency
Dim LOC_CANCELA As Integer
Dim pasa  As Integer
Dim LOC_CTA_CLI(2) As String * 12
Dim LOC_DES_CLI(2) As String * 50
Dim PSART_RELA As rdoQuery
Dim art_rela As rdoResultset

Public Sub BLOQUEA_TEXT(Optional o1, Optional o2, Optional o3, Optional o4, Optional o5, Optional o6, Optional o7, Optional o8, Optional o9, Optional o10)
'** BLOQUEA TEXTBOX  CANTIDAD DE OBJECTOS **
If Not IsMissing(o1) Then
 o1.Enabled = False
End If
If Not IsMissing(o2) Then
 o2.Enabled = False
End If
If Not IsMissing(o3) Then
 o3.Enabled = False
End If
If Not IsMissing(o4) Then
 o4.Enabled = False
End If
If Not IsMissing(o5) Then
 o5.Enabled = False
End If
If Not IsMissing(o6) Then
 o6.Enabled = False
End If
If Not IsMissing(o7) Then
 o7.Enabled = False
End If
If Not IsMissing(o8) Then
 o8.Enabled = False
End If
If Not IsMissing(o9) Then
 o9.Enabled = False
End If
If Not IsMissing(o10) Then
 o10.Enabled = False
End If
End Sub
Public Sub DESBLOQUEA_TEXT(Optional o1, Optional o2, Optional o3, Optional o4, Optional o5, Optional o6, Optional o7, Optional o8, Optional o9, Optional o10)
'** BLOQUEA TEXTBOX  CANTIDAD DE OBJECTOS **
If Not IsMissing(o1) Then
 o1.Enabled = True
End If
If Not IsMissing(o2) Then
 o2.Enabled = True
End If
If Not IsMissing(o3) Then
 o3.Enabled = True
End If
If Not IsMissing(o4) Then
 o4.Enabled = True
End If
If Not IsMissing(o5) Then
 o5.Enabled = True
End If
If Not IsMissing(o6) Then
 o6.Enabled = True
End If
If Not IsMissing(o7) Then
 o7.Enabled = True
End If
If Not IsMissing(o8) Then
 o8.Enabled = True
End If
If Not IsMissing(o9) Then
 o9.Enabled = True
End If
If Not IsMissing(o10) Then
 o10.Enabled = True
End If
End Sub

Private Sub art_codpro_GotFocus()
frmARTI.F14.Visible = False
End Sub

Private Sub art_codpro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If DS.Enabled Then
    DS.SetFocus
    SendKeys "%{UP}"
  Else
    DS_KeyPress 13
  End If
End If
End Sub

Private Sub art_codpro_LostFocus()
Fvarios.Refresh
End Sub

Private Sub art_familia_Click()
 art_subfam.Clear
 art_grupo.Clear
 art_numero.Clear
' art_linea.Clear
' art_marca.Clear

End Sub

Private Sub art_familia_GotFocus()
frmARTI.F14.Visible = False
End Sub

Private Sub art_familia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If loc_tipo = "V" Then
        DoEvents
        art_subfam.SetFocus
        DoEvents
        SendKeys "%{UP}"
        Exit Sub
    End If
End If
End Sub

Private Sub art_familia_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos
If KeyCode <> 45 Then
 Exit Sub
End If
Flag_Bloq = "A"
wpos = art_familia.ListIndex
PUB_TIPREG = Mid(art_familia.ToolTipText, 13, Len(art_familia.ToolTipText))
PUB_CODCIA = LK_CODCIA
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If
Load FrmDatArti
FrmDatArti.Caption = "FAMILIAS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
Flag_Bloq = ""
DoEvents
LLENADO_FAM
DoEvents
On Error GoTo sigue
art_familia.ListIndex = wpos
On Error GoTo 0
art_familia.SetFocus
SendKeys "%{up}"
Exit Sub
sigue:
Resume Next
End Sub

Private Sub art_familia_LostFocus()
txtnombre.Text = ARMA_NOMBRE
Dim wpos As Integer
Dim WFAMI2 As Integer
If Flag_Bloq = "A" Then
 Exit Sub
End If
If Trim(art_familia.Text) = "" Then
 art_subfam.Clear
 Exit Sub
End If
wpos = art_subfam.ListIndex
WFAMI2 = Val(Trim(Right(art_familia.Text, 6)))
PUB_TIPREG = 123
LLENADO_SUBFAM art_subfam, WFAMI2
Fvarios.Refresh
On Error GoTo sigue
art_subfam.ListIndex = wpos
Exit Sub
sigue:
Resume Next
End Sub

Private Sub art_grupo_Click()
 art_numero.Clear
 'art_linea.Clear
 'art_marca.Clear
End Sub

Private Sub art_grupo_GotFocus()
frmARTI.F14.Visible = False
End Sub

Private Sub art_grupo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 art_numero.SetFocus
 SendKeys "%{up}"
End If
End Sub

Private Sub art_grupo_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wfami As Integer
If KeyCode <> 45 Then
 Exit Sub
End If
On Error GoTo sigue
Dim wpos
Pres = "INSERT"
wpos = art_grupo.ListIndex
wfami = Val(Trim(Right(art_subfam.Text, 6)))
PUB_CODART = wfami
PUB_TIPREG = Mid(art_grupo.ToolTipText, 13, Len(art_grupo.ToolTipText))
PUB_CODCIA = LK_CODCIA
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If

Load FrmDatArti
FrmDatArti.Caption = "GRUPOS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents

LLENADO_SUBFAM art_grupo, wfami

art_grupo.ListIndex = wpos
On Error GoTo 0
art_grupo.SetFocus
SendKeys "%{up}"
Exit Sub
sigue:
Resume Next
End Sub

Private Sub art_grupo_LostFocus()
Fvarios.Refresh
txtnombre.Text = ARMA_NOMBRE
If Pres = "INSERT" Then
 Pres = ""
 Exit Sub
End If

  If Trim(art_grupo.Text) = "" Then
   art_numero.Clear
   Exit Sub
  End If
  wpos = art_numero.ListIndex
  WFAMI2 = Val(Trim(Right(art_grupo.Text, 6)))
  PUB_TIPREG = 130
  LLENADO_SUBFAM art_numero, WFAMI2
  Fvarios.Refresh
  On Error GoTo sigue
  art_numero.ListIndex = wpos
  Exit Sub

sigue:
Resume Next
End Sub


Private Sub art_linea_Click()
  'art_marca.Clear

End Sub

Private Sub art_linea_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wSubGrupo As Integer
If KeyCode <> 45 Then
 Exit Sub
End If
Dim wpos
Pres = "INSERT"
wpos = art_linea.ListIndex
wSubGrupo = Val(Trim(Right(art_numero.Text, 6)))
'**************************
'ojo
'**************************
If Not art_linea.ToolTipText = "" Then
 PUB_TIPREG = Mid(art_linea.ToolTipText, 13, Len(art_linea.ToolTipText))
End If
PUB_CODCIA = LK_CODCIA
PUB_CODART = wSubGrupo
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If
Load FrmDatArti
FrmDatArti.Caption = "GRUPOS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENADO_SUBFAM art_linea, wSubGrupo
On Error GoTo sigue
art_linea.ListIndex = wpos
On Error GoTo 0
art_linea.SetFocus
SendKeys "%{up}"
Exit Sub
sigue:
Resume Next


End Sub

Private Sub art_Linea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 art_marca.SetFocus
 SendKeys "%{up}"
End If

End Sub

Private Sub art_Linea_LostFocus()
txtnombre.Text = ARMA_NOMBRE
If Pres = "INSERT" Then
 Pres = ""
 Exit Sub
End If
  If Trim(art_linea.Text) = "" Then
   art_marca.Clear
   Exit Sub
  End If
  wpos = art_marca.ListIndex
  WFAMI2 = Val(Trim(Right(art_linea.Text, 6)))
  PUB_TIPREG = 132
  LLENADO_SUBFAM art_marca, WFAMI2
  Fvarios.Refresh
  On Error GoTo sigue
  art_marca.ListIndex = wpos
  Exit Sub
sigue:
Resume Next
End Sub

Private Sub art_marca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 art_codpro.SetFocus
 SendKeys "%{up}"
End If

End Sub
Private Sub art_marca_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
  OP_FORM = "C"
  Load frmprecol
  frmprecol.Show 1
  OP_FORM = ""
  Exit Sub
End If
If KeyCode <> 45 Then
 Exit Sub
End If
Dim wpos
wpos = art_marca.ListIndex
wfami = Val(Trim(Right(art_linea.Text, 6)))
PUB_TIPREG = Mid(art_marca.ToolTipText, 13, Len(art_marca.ToolTipText))
PUB_CODCIA = LK_CODCIA
PUB_CODART = wfami
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If
Load FrmDatArti
FrmDatArti.Caption = "Marca  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENADO_SUBFAM art_marca, wfami
On Error GoTo sigue
art_marca.ListIndex = wpos
On Error GoTo 0
art_marca.SetFocus
SendKeys "%{up}"
Exit Sub
sigue:
Resume Next

End Sub

Private Sub art_marca_LostFocus()
txtnombre.Text = ARMA_NOMBRE
End Sub

Private Sub art_numero_Click()
If Not art_numero.ListIndex = art_numero.ListIndex Then
 'art_linea.Clear
 'art_marca.Clear
End If
End Sub

Private Sub art_numero_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wgrupo As Integer
If KeyCode = 116 Then
  OP_FORM = "T"
  Load frmprecol
  frmprecol.Show 1
  OP_FORM = ""
  Exit Sub
End If
If KeyCode <> 45 Then
 Exit Sub
End If
Pres = "INSERT"
Dim wpos

wpos = art_numero.ListIndex
wgrupo = Val(Trim(Right(art_grupo.Text, 6)))
PUB_TIPREG = Mid(art_numero.ToolTipText, 13, Len(art_numero.ToolTipText))
PUB_CODCIA = LK_CODCIA
PUB_CODART = wgrupo
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If
Load FrmDatArti
FrmDatArti.Caption = "GRUPOS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENADO_SUBFAM art_numero, wgrupo
On Error GoTo sigue
art_numero.ListIndex = wpos
On Error GoTo 0
art_numero.SetFocus
SendKeys "%{up}"
Exit Sub
sigue:
Resume Next

End Sub

Private Sub art_numero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 art_linea.SetFocus
 SendKeys "%{up}"
End If

End Sub

Private Sub art_numero_LostFocus()
txtnombre.Text = ARMA_NOMBRE
If Pres = "INSERT" Then
 Pres = ""
 Exit Sub
End If

  If Trim(art_numero.Text) = "" Then
   'art_linea.Clear
   Exit Sub
  End If
  wpos = art_linea.ListIndex
  WFAMI2 = Val(Trim(Right(art_numero.Text, 6)))
  PUB_TIPREG = 131
  LLENADO_SUBFAM art_linea, WFAMI2
  Fvarios.Refresh
  On Error GoTo sigue
  art_linea.ListIndex = wpos
  Exit Sub
sigue:
Resume Next
End Sub

Private Sub art_plancha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 art_codpro.SetFocus
 SendKeys "%{up}"
End If
End Sub

Private Sub art_plancha_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 45 Then
 Exit Sub
End If
Dim wpos
wpos = art_plancha.ListIndex
PUB_TIPREG = Mid(art_plancha.ToolTipText, 13, Len(art_plancha.ToolTipText))
PUB_CODCIA = LK_CODCIA
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If
Load FrmDatArti
FrmDatArti.Caption = "Lote  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENADO_PLANCHA
On Error GoTo sigue
art_plancha.ListIndex = wpos
On Error GoTo 0
art_plancha.SetFocus
SendKeys "%{up}"
Exit Sub
sigue:

End Sub

Private Sub art_plancha_LostFocus()
txtnombre.Text = ARMA_NOMBRE
End Sub

Private Sub art_subfam_Click()
art_grupo.Clear
art_numero.Clear
'art_linea.Clear
'art_marca.Clear

End Sub

Private Sub art_subfam_GotFocus()
frmARTI.F14.Visible = False
End Sub
Private Sub art_subfam_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   art_grupo.SetFocus
   SendKeys "%{up}"
   Exit Sub
End If
End Sub
Private Sub art_subfam_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wfami As Integer
If KeyCode <> 45 Then
 Exit Sub
End If
wfami = Val(Trim(Right(art_familia.Text, 6)))
If Mid(art_subfam.ToolTipText, 13, Len(art_subfam.ToolTipText)) = "" Then
  Exit Sub
End If
Pres = "INSERT"
'*******************
Dim wpos
wpos = art_subfam.ListIndex
PUB_TIPREG = Val(Mid(art_subfam.ToolTipText, 13, Len(art_subfam.ToolTipText)))
PUB_CODCIA = LK_CODCIA
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If
PUB_CODART = wfami
Load FrmDatArti
FrmDatArti.Caption = "SUB-FAMILIAS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENADO_SUBFAM art_subfam, wfami
On Error GoTo sigue
art_subfam.ListIndex = wpos
On Error GoTo 0
art_subfam.SetFocus
SendKeys "%{up}"
Exit Sub
sigue:
Resume Next

End Sub

Private Sub art_subfam_LostFocus()

Fvarios.Refresh
txtnombre.Text = ARMA_NOMBRE
If Pres = "INSERT" Then
 Pres = ""
 Exit Sub
End If
  If Trim(art_subfam.Text) = "" Then
   art_grupo.Clear
   Exit Sub
  End If
  wpos = art_grupo.ListIndex
  WFAMI2 = Val(Trim(Right(art_subfam.Text, 6)))
  PUB_TIPREG = 129
  LLENADO_SUBFAM art_grupo, WFAMI2
  Fvarios.Refresh
  On Error GoTo sigue
  art_grupo.ListIndex = wpos
  Exit Sub

Exit Sub
sigue:
Resume Next
End Sub

Private Sub CERO_Click()
PUB_KEY = 0
PUB_CODCIA = LK_CODCIA
LOC_OPER = 2
LEER_LOC
If artloc_key.EOF Then
  ARTI_CERO
  MsgBox "Producto Creado ..", 48, Pub_Titulo
End If
End Sub

Private Sub cmbcal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAdd_Click
End Sub

Private Sub CmbCalidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  art_codpro.SetFocus
  DoEvents
  SendKeys "%{up}"
End If

End Sub

Private Sub CmbCalidad_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 45 Then
 Exit Sub
End If
Dim wpos
wpos = CmbCalidad.ListIndex
PUB_TIPREG = Mid(CmbCalidad.ToolTipText, 13, Len(CmbCalidad.ToolTipText))
PUB_CODCIA = LK_CODCIA
Load FrmDatArti
FrmDatArti.Caption = "CALIDAD  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENADO_CAL
On Error GoTo sigue
CmbCalidad.ListIndex = wpos
On Error GoTo 0
'CmbCalidad.SetFocus
'SendKeys "%{up}"
Exit Sub
sigue:
Resume Next

End Sub

Private Sub CmbCalidad_LostFocus()
Fvarios.Refresh
End Sub

Private Sub cmdAdd_Click()
Dim wnombre
Dim wrellena As String
Dim WART_COSPRO As Currency
Dim WART_COSPRO_ANT As Currency
Dim WART_COSTO_ULT As Currency

If Trim(cmbcal.Text) = "" Then
 MsgBox " Seleccione su Calidad.", 48, Pub_Titulo
 cmbcal.SetFocus
 Exit Sub
End If
If Left(cmbcal.Text, 1) = "<" Then
 MsgBox "No Existe mas Calidades ", 48, Pub_Titulo
 cmbcal.SetFocus
 Exit Sub
End If

wnombre = InputBox("Ingrese la Descripción del Articulo :", Pub_Titulo, Trim(txtnombre.Text))
If wnombre = "" Then
  Screen.MousePointer = 0
  Exit Sub
End If
If Trim(wnombre) = "" Then
  Screen.MousePointer = 0
  MsgBox "Descripción NO Validad.", 48, Pub_Titulo
  Exit Sub
End If
LOC_NOMBRE = wnombre
VAR_NEWCAL = 1
LOC_CALIDAD = Val(Right(cmbcal.Text, 3))
wrellena = String(LOC_CALIDAD - 1, "*")
LOC_ALTERNO = Trim(txt_alterno.Text) + wrellena
pu_codcia = LK_CODCIA
PUB_CODART = artloc_llave!ART_KEY
SQ_OPER = 1
LEER_ARM_LLAVE
WART_COSPRO = arm_llave!arm_cospro
WART_COSTO_ULT = arm_llave!ARM_COSTO_ULT

On Error GoTo ESCAPA
CN.Execute "Begin Transaction", rdExecDirect
pub_cadena = "SELECT * FROM CONTROLL"
Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
LOC_ORIGINAL = GENERA_CODI()
LOC_CODART2 = LOC_ORIGINAL
GRABAR_ARTI
PUB_CODART = LOC_ORIGINAL
SQ_OPER = 1
LEER_ARM_LLAVE
If arm_llave.EOF Then
   Screen.MousePointer = 0
   arm_llave.AddNew
   arm_llave!ARM_CODART = LOC_ORIGINAL
   arm_llave!ARM_CODCIA = LK_CODCIA
   arm_llave!ARM_STOCK = 0
   arm_llave!ARM_INGRESOS = 0
   arm_llave!ARM_SALIDAS = 0
   arm_llave!arm_cospro = WART_COSPRO
   arm_llave!arm_stock2 = 0
   arm_llave!ARM_Saldo_n = 0
   arm_llave!ARM_SALDO_N2 = 0
   arm_llave!ARM_saldo_s = 0
   arm_llave!arm_saldo_s2 = 0
   arm_llave!ARM_COSTO_ULT = WART_COSTO_ULT
   arm_llave.Update
Else
  MsgBox "Codigo Existe en tabla: Articulo verificar ...", 48, Pub_Titulo
  GoTo ESCAPA
End If
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
   pu_alterno = Trim(txt_alterno.Text)
Else
   PUB_KEY = Val(Txt_key.Text)
End If
PUB_CODCIA = LK_CODCIA
LOC_OPER = 1
LEER_LOC
If artloc_llave.EOF Then
  GoTo ESCAPA
 Exit Sub
End If

artloc_llave.Edit
artloc_llave!ART_CODART2 = LOC_CODART2
'artloc_llave!ART_COSPRO = Val(tcospro.text)
artloc_llave.Update
gridrel.TextMatrix(1, 0) = LOC_ORIGINAL
gridrel.TextMatrix(1, 1) = LOC_ALTERNO
gridrel.TextMatrix(1, 2) = LOC_NOMBRE
gridrel.TextMatrix(1, 3) = Left(cmbcal.Text, 40)
txtcodigo2.Text = LOC_CODART2
cmbcal.Visible = False
lblcal.Visible = False
cmdAdd.Visible = False
cmdquitar.Visible = True
con_llave.Close
CN.Execute "Commit Transaction", rdExecDirect
VAR_NEWCAL = 0
On Error GoTo 0
Exit Sub
ESCAPA:
VAR_NEWCAL = 0
If con_llave Is Nothing Then
 con_llave.Close
 CN.Execute "Rollback Transaction", rdExecDirect
End If
MsgBox " Intente Nuevamente..", 48, Pub_Titulo

     
End Sub

Private Sub cmdagregar_Click()
If Trim(Txt_key.Text) = "1" Then
     MENSAJE_ARTI "No Procede. .."
     Exit Sub
End If
If Left(cmdAgregar.Caption, 2) = "&A" Then
    cmdAgregar.Caption = "&Grabar"
    cmdCancelar.Enabled = True
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    LIMPIA_ARTI
    frmARTI.decimales.ListIndex = 1
    If frmARTI.CmbCalidad.ListCount <> 0 Then
       frmARTI.CmbCalidad.ListIndex = 0
    End If
    If frmARTI.art_grupo.ListCount <> 0 Then
       frmARTI.art_grupo.ListIndex = 0
    End If
    frmARTI.Txt_key = GENERA_CODI
    DESBLOQUEA_TEXT txtnombre, decimales, DS, txtcospro, art_situacion, art_linea, art_numero, art_marca, art_plancha
    DESBLOQUEA_TEXT art_grupo, art_familia, art_subfam, grid_unid, txtMin, txtMax, art_codpro, txtcodigo2
    DESBLOQUEA_TEXT cheservi(0), cheservi(1), cheservi(2), exigv, txtcospro, cmddolares, txtpeso, txtfechault, checambio, txtlitro
    DESBLOQUEA_TEXT txtcolor, txtmarca
    If LK_EMP = "HER" Then
     DESBLOQUEA_TEXT txtpor1, txtpor2, txtpor3, txtpor4, txtpor5, txtpor6
     picfoto.Visible = True
     End If
    BLOQUEA_TEXT Txt_key
    If LK_MONEDA = "D" Then
        frmARTI.DS.ListIndex = 1
        frmARTI.DS.Enabled = False
    ElseIf LK_MONEDA = "S" Then
        frmARTI.DS.ListIndex = 0
        frmARTI.DS.Enabled = False
    ElseIf LK_MONEDA = "A" Then
        frmARTI.DS.ListIndex = -1
    End If
    Flag_Inicial = "A"
    CABEZA_UNID
    'grid_unid.Rows = grid_unid.Rows + 1
    grid_unid.RowHeight(1) = 285
    If LK_EMP = "PIU" Then
       grid_unid.TextMatrix(1, 0) = "PARES"
    ElseIf LK_EMP = "CAM" Then
       grid_unid.TextMatrix(1, 0) = "Kg."
    Else
       grid_unid.TextMatrix(1, 0) = "UNIDAD"
    End If
    grid_unid.TextMatrix(1, 1) = "1.00"
    grid_unid.TextMatrix(1, 3) = "0.0000"
    grid_unid.COL = 4
    grid_unid.CellForeColor = QBColor(9)
    grid_unid.TextMatrix(1, 4) = "0.00"
    grid_unid.TextMatrix(1, 5) = "0.00"
    grid_unid.COL = 6
    grid_unid.CellForeColor = QBColor(9)
    grid_unid.TextMatrix(1, 6) = "0.00"
    grid_unid.TextMatrix(1, 7) = "0.00"
    grid_unid.COL = 8
    grid_unid.CellForeColor = QBColor(9)
    grid_unid.TextMatrix(1, 8) = "0.00"
    grid_unid.TextMatrix(1, 9) = "0.00"
    grid_unid.COL = 10
    grid_unid.CellForeColor = QBColor(9)
    grid_unid.TextMatrix(1, 10) = "0.00"
    grid_unid.TextMatrix(1, 11) = "0.00"
    grid_unid.COL = 12
    grid_unid.CellForeColor = QBColor(9)
    grid_unid.TextMatrix(1, 12) = "0.00"
    grid_unid.TextMatrix(1, 13) = "0.00"
    grid_unid.TextMatrix(1, 14) = "A"
''*********************************Precio6???
    grid_unid.COL = 29
    grid_unid.CellForeColor = QBColor(9)
    grid_unid.TextMatrix(1, 29) = "0.00"
    grid_unid.TextMatrix(1, 30) = "0.00"
''*****************************************
    Flag_Inicial = ""
    grid_unid.COL = 0
    lblUnidad.Caption = "UNIDAD"
    DESBLOQUEA_TEXT txt_alterno
    pasa = 1
    If txt_alterno.Visible And txt_alterno.Enabled Then
      txt_alterno.SetFocus
    ElseIf txtnombre.Visible And txtnombre.Enabled Then
     txtnombre.SetFocus
    End If
    cheservi(0).Value = True
    MANOS(0).Enabled = False
    MANOS(1).Enabled = False
    fcomun.Refresh
    Fvarios.Refresh
Else 'Actualiza Datos
    If frmARTI.DS.ListIndex = -1 Then
       MsgBox "Determinar la Moneda del Articulo ..", 48, Pub_Titulo
       frmARTI.DS.SetFocus
       SendKeys "%{UP}"
       Exit Sub
    End If
    If Trim(frmARTI.CmbCalidad.Text) = "" Then
       MsgBox "Definir Calidad en,  Tablas del Sistema ", 48, Pub_Titulo
       Exit Sub
    End If
    If Not CONSIS_ARTI Then
       Exit Sub
    End If
    If Not CONSIS_UNIDAD Then
       Screen.MousePointer = 0
       MsgBox "Verificar Datos de unidad no valido..", 48, Pub_Titulo
       grid_unid.SetFocus
       Exit Sub
    End If
    If LK_FLAG_ORIGINAL <> "A" Then
       If Trim(txt_alterno.Text) = "" Then
          MsgBox "Codigo alterno no valido ..!!! Verificar ", 48, Pub_Titulo
          Azul frmARTI.txt_alterno, frmARTI.txt_alterno
          GoTo fin
       End If
       SQ_OPER = 3
       pu_alterno = txt_alterno.Text
       pu_codcia = LK_CODCIA
       LEER_ART_LLAVE
       'If Not art_llave_alt.EOF Then   ' quitado gts para grabar codigos alternos iguales
'bloqeuado por mic luego desbloquear
        '    Debug.Print txt_alterno.Text & "  *  " & txtnombre.Text
        ' MsgBox "Codigo Alterno  EXISTE ...!!! Verificar " & Chr(13) & Trim(art_llave_alt!art_alterno) & " : " & Trim(art_llave_alt!art_nombre), 48, Pub_Titulo
        '  Azul frmARTI.txt_alterno, frmARTI.txt_alterno
        '  GoTo fin
       'End If
    End If
    If pasa = 1 Then
       If EXISTE_ART(txtnombre.Text, Trim(Txt_key.Text)) Then
          MENSAJE_ARTI "Existen algunos Articulos con estos NOMBRES .."
          frmARTI.ListExiste.SetFocus
          'Exit Sub
       End If
    End If
    pasa = 0
     Screen.MousePointer = 11
'     On Error GoTo ESCAPA
     pub_cadena = "SELECT * FROM CONTROLL"
     Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
     CN.Execute "Begin Transaction", rdExecDirect
     pub_cadena = "SELECT * FROM CONTROLL"
     Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
     frmARTI.Txt_key = GENERA_CODI
     PUB_KEY = Val(frmARTI.Txt_key)
     If Trim(Nulo_Valors(par_llave!par_art_cias)) <> "" Then
        xcuenta = 1
        For fila = 1 To 30
          pu_codcia = Mid(Trim(par_llave!par_art_cias), xcuenta, 2)
'          If pu_codcia = "05" And LK_EMP = "CAM" Then GoTo SALE
          If Trim(pu_codcia) = "" Then Exit For
             GRABAR_ARTI
            'GoSub IR_POR_CIA
SALE:
          xcuenta = xcuenta + 2
        Next fila
     Else
        GRABAR_ARTI
     End If
     con_llave.Close
     CN.Execute "Commit Transaction", rdExecDirect
     On Error GoTo 0
     cmdAgregar.Caption = "&Adicionar"
     'cmdCancelar.Enabled = True
     cmdEliminar.Enabled = True
     cmdModificar.Enabled = True
     LIMPIA_ARTI
     BLOQUEA_TEXT txtnombre, CmbCalidad, decimales, DS, txtcospro, art_situacion, art_linea, art_numero, art_marca, art_plancha
     BLOQUEA_TEXT art_grupo, art_familia, art_subfam, grid_unid, txtMin, txtMax, art_codpro, txtcodigo2
     BLOQUEA_TEXT cheservi(0), cheservi(1), cheservi(2), exigv, txtcospro, cmddolares, txtpeso, txtfechault, checambio, txtlitro
     BLOQUEA_TEXT txtcolor, txtmarca
     If LK_EMP = "HER" Then
       BLOQUEA_TEXT txtpor1, txtpor2, txtpor3, txtpor4, txtpor5, txtpor6
     End If
     If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       DESBLOQUEA_TEXT txt_alterno
       BLOQUEA_TEXT Txt_key
       txt_alterno.SetFocus
     Else
       DESBLOQUEA_TEXT Txt_key
       BLOQUEA_TEXT txt_alterno
       Txt_key.SetFocus
     End If
     Screen.MousePointer = 0
     MANOS(0).Enabled = True
     MANOS(1).Enabled = True
     MENSAJE_ARTI "Articulo   AGREGADO ... "
End If
Exit Sub
ESCAPA:
'    If con_llave Is Nothing Then
     con_llave.Close
     CN.Execute "Rollback Transaction", rdExecDirect
 '   End If
    If Err.Number = 40002 Then
        MsgBox "Hay Error en la LLave ..Intente Nuevamente. "
    ElseIf Err.Number <> 0 Then
        MsgBox Err.Number & "  " & Err.Description & "  Intente Nuevamente."
    End If
    cmdAgregar.Caption = "&Adicionar"
    cmdCancelar.Enabled = True
    cmdEliminar.Enabled = True
    cmdModificar.Enabled = True
    LIMPIA_ARTI
    BLOQUEA_TEXT txtnombre, CmbCalidad, decimales, DS, txtcospro, art_situacion, art_linea, art_numero, art_marca, art_plancha
    BLOQUEA_TEXT art_grupo, art_familia, art_subfam, grid_unid, txtMin, txtMax, art_codpro, txtcodigo2
    BLOQUEA_TEXT cheservi(0), cheservi(1), cheservi(2), exigv, txtcospro, cmddolares, txtpeso, txtfechault, checambio, txtlitro
    BLOQUEA_TEXT txtcolor, txtmarca
    If LK_EMP = "HER" Then
     BLOQUEA_TEXT txtpor1, txtpor2, txtpor3, txtpor4, txtpor5, txtpor6
    End If
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
     DESBLOQUEA_TEXT txt_alterno
     txt_alterno.SetFocus
    Else
     DESBLOQUEA_TEXT Txt_key
     Txt_key.SetFocus
    End If
    Screen.MousePointer = 0
    MANOS(0).Enabled = True
    MANOS(1).Enabled = True
   Exit Sub
fin:
End Sub

Private Sub cmdagregar_GotFocus()
If ListView1.Visible Then
  ListView1.Visible = False
  Txt_key.Text = ""
End If
End Sub

Private Sub cmdAgregar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
'    frmARTI.txt_key.SetFocus
End If

End Sub


Private Sub cmdcancelar_Click()
PROCESO_CANCELAR
End Sub

Private Sub cmdCancelar_GotFocus()
If ListView1.Visible Then
  ListView1.Visible = False
  Txt_key.Text = ""
End If
End Sub

Private Sub cmdCerrar_Click()
cmdcancelar_Click
frmARTI.Hide
End Sub

Private Sub cmdCerrar_GotFocus()
If ListView1.Visible Then
  ListView1.Visible = False
  Txt_key.Text = ""
End If
End Sub

Private Sub cmdconfirma_Click()
Dim cnPre As New ADODB.Connection
Dim cod As Long
Dim i As Long
Dim fam As Integer
Dim SubFam As Integer
Dim Nombre As String
Dim marca As String
     pasa = 0
     updatePrecios = True
     If Left(cmdModificar.Caption, 2) = "&G" Then
       'TOMA VALORES
       cod = Val(Me.Txt_key.Text)
       fam = Val(Right(art_familia.Text, 5))
       SubFam = Val(Right(art_subfam.Text, 5))
       Nombre = txtnombre.Text
       marca = txtmarca.Text
       CmdModificar_Click
       MsgBox "Confirma Grabacion de TODOS los Articulos Seleccionados?"
       If Op(0).Value Then
        pgb_Progress.Value = 0
        pgb_Progress.max = Me.ListExiste.Rows
        pgb_Progress.Visible = True
        
        'ONLY MOVE
        cnPre.Open "Provider=SQLOLEDB.1;Data Source=" & CONST_SERVER & ";Initial Catalog=BDATOS;User ID=" & CONST_UID & ";Password=" & CONST_PWD
        'cnPre.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=sa;Initial Catalog=bdatos;Data Source=laptop;Password=" & CONST_PWD
        
        On Error GoTo anulaTodo
        cnPre.BeginTrans
        For i = 1 To Me.ListExiste.Rows - 1
         If Trim(Me.ListExiste.TextMatrix(i, 1)) = "" Then
          GoTo nextArti
         End If
         If ListExiste.TextMatrix(i, 4) = "0" Then
          GoTo nextArti
         End If
         
         cnPre.Execute "EXEC CopyPrecios '" & PUB_CODCIA & "', " & cod & ", " & Val(Trim(Me.ListExiste.TextMatrix(i, 1))) & "", , adExecuteNoRecords
         cnPre.Execute "EXEC UpdateArticulo '" & PUB_CODCIA & "', " & Val(Trim(Me.ListExiste.TextMatrix(i, 1))) & ",'" & Nombre & "', " & fam & ", " & SubFam & ", '" & marca & "' ", , adExecuteNoRecords
         
nextArti:
         pgb_Progress.Value = pgb_Progress.Value + 1
        Next i
        cnPre.CommitTrans
        GoTo normal
anulaTodo:

cnPre.RollbackTrans
MsgBox Err.Description, vbExclamation, "Precios"

normal:
        pgb_Progress.Visible = False
        frmARTI.F14.Visible = False
        updatePrecios = False
        cmdcancelar_Click
       Else
        cmdcancelar_Click
        updatePrecios = False
        frmARTI.F14.Visible = False
       End If
      Else
       cmdagregar_Click
     End If
SALIR:
     Set cnPre = Nothing
End Sub


Private Sub cmddolares_Click()
Dim WSPOR As Currency
If LK_EMP = "3AA" Then
  If cmddolares.Tag = "D" Then
    cmddolares.Caption = "Lista de Precios en S/. (Nuevos Soles)"
    cmddolares.Tag = "S"
  Else
    cmddolares.Caption = "Lista de Precios en US$. (Dolares Americanos)"
    cmddolares.Tag = "D"
  End If
Else
  If cmddolares.Tag = "D" Then
    cmddolares.Caption = "Lista de Precios en S/. (Nuevos Soles)"
    cmddolares.Tag = "S"
  Else
    cmddolares.Caption = "Lista de Precios en US$. (Dolares Americanos)"
    cmddolares.Tag = "D"
  End If
End If
For fila = 1 To grid_unid.Rows - 1
    If cmddolares.Tag = "D" Then
     grid_unid.TextMatrix(fila, 5) = grid_unid.TextMatrix(fila, 16)
     grid_unid.TextMatrix(fila, 7) = grid_unid.TextMatrix(fila, 17)
     grid_unid.TextMatrix(fila, 9) = grid_unid.TextMatrix(fila, 18)
     grid_unid.TextMatrix(fila, 11) = grid_unid.TextMatrix(fila, 19)
     grid_unid.TextMatrix(fila, 13) = grid_unid.TextMatrix(fila, 20)
''*************************************PRECIO6??
     grid_unid.TextMatrix(fila, 30) = grid_unid.TextMatrix(fila, 31)
''*************************************
    Else
     grid_unid.TextMatrix(fila, 5) = grid_unid.TextMatrix(fila, 21)
     grid_unid.TextMatrix(fila, 7) = grid_unid.TextMatrix(fila, 22)
     grid_unid.TextMatrix(fila, 9) = grid_unid.TextMatrix(fila, 23)
     grid_unid.TextMatrix(fila, 11) = grid_unid.TextMatrix(fila, 24)
     grid_unid.TextMatrix(fila, 13) = grid_unid.TextMatrix(fila, 25)
''*************************************PRECIO6??
     grid_unid.TextMatrix(fila, 30) = grid_unid.TextMatrix(fila, 32)
''*************************************
    End If
    If cmddolares.Tag = "D" Then
         grid_unid.TextMatrix(fila, 3) = redondea(Val(grid_unid.TextMatrix(fila, 27)) / LK_TIPO_CAMBIO)
     Else
         grid_unid.TextMatrix(fila, 3) = redondea(Val(grid_unid.TextMatrix(fila, 27)))
     End If
    
     If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then
       WSPOR = (Val(grid_unid.TextMatrix(fila, 5)) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
     End If
     grid_unid.TextMatrix(fila, 4) = Format(WSPOR, "0.00")
     If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then
       WSPOR = (Val(grid_unid.TextMatrix(fila, 7)) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
     End If
     grid_unid.TextMatrix(fila, 6) = Format(WSPOR, "0.00")
     If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then
       WSPOR = (Val(grid_unid.TextMatrix(fila, 9)) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
     End If
     grid_unid.TextMatrix(fila, 8) = Format(WSPOR, "0.00")
     If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then
       WSPOR = (Val(grid_unid.TextMatrix(fila, 11)) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
     End If
     grid_unid.TextMatrix(fila, 10) = Format(WSPOR, "0.00")
     If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then
       WSPOR = (Val(grid_unid.TextMatrix(fila, 13)) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
     End If
     grid_unid.TextMatrix(fila, 12) = Format(WSPOR, "0.00")
''****************************************PRECIO6
     If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then
       WSPOR = (Val(grid_unid.TextMatrix(fila, 30)) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
     End If
     grid_unid.TextMatrix(fila, 29) = Format(WSPOR, "0.00")
''*************************************************
     
    
 Next fila
'If Trim(DS.Text) = "S" Then
'   If cmddolares.Tag = "D" Then
'     grid_unid.Enabled = False
'   Else
     grid_unid.Enabled = True
     If grid_unid.Enabled Then grid_unid.SetFocus
'   End If
'Else
'   If cmddolares.Tag = "D" Then
'    grid_unid.Enabled = True
'    grid_unid.SetFocus
'   Else
'    grid_unid.Enabled = False
'   End If
'End If


End Sub

Private Sub cmdEliminar_Click()
Dim ws_codcia As String
Dim WS_CODART As Currency
Dim flag_puntos As String * 1
On Error GoTo ESCAPA
If Len(Txt_key) = 0 Or Len(txtnombre.Text) = 0 Then
    If Not Trim(Txt_key) = "1" Then
       Screen.MousePointer = 0
       MENSAJE_ARTI "NO a seleccionado ningun Articulo... !"
'       txt_key.SetFocus
       Exit Sub
    End If
End If
  Dim PS_REP01 As rdoQuery
  Dim llave_rep01 As rdoResultset
  Dim OpenForms
  
  WS_CODART = artloc_llave!ART_KEY
  If Trim(GEN!gen_ART_CIAS) <> "" Then
        xcuenta = 1
        For fila = 1 To 30
            ws_codcia = Mid(Trim(GEN!gen_ART_CIAS), xcuenta, 2)
            If Trim(ws_codcia) = "" Then Exit For
            SQ_OPER = 1
            PUB_CODART = WS_CODART
            pu_codcia = ws_codcia
            LEER_ARM_LLAVE
            If Not arm_llave.EOF Then
                If arm_llave!ARM_STOCK = 0 And arm_llave!ARM_STOCK = 0 And arm_llave!ARM_STOCK = 0 Then
                Else
                    LblMensaje.Visible = False
                    Screen.MousePointer = 0
                    MsgBox "NO se Puede Eliminar ...  ARTICULO CON HISTORIA ", 48, Pub_Titulo
                    Exit Sub
                End If
            End If
            xcuenta = xcuenta + 2
        Next fila
  End If
  
  Screen.MousePointer = 11
  LblMensaje.Visible = True
  LblMensaje.Caption = "Verificando Data.  un Momento..."
  WS_CODART = artloc_llave!ART_KEY
  SQ_OPER = 1
  PUB_CODART = artloc_llave!ART_KEY
  pu_codcia = LK_CODCIA
  LEER_ARM_LLAVE
  If Not arm_llave.EOF Then
      If arm_llave!ARM_STOCK = 0 And arm_llave!ARM_STOCK = 0 And arm_llave!ARM_STOCK = 0 Then
      Else
          LblMensaje.Visible = False
          Screen.MousePointer = 0
          MsgBox "NO se Puede Eliminar ...  ARTICULO CON HISTORIA ", 48, Pub_Titulo
          Exit Sub
      End If
  End If
  pub_cadena = "SELECT FAR_CODART FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODART = ?  AND FAR_ESTADO <> 'E' "
  Set PS_REP01 = CN.CreateQuery("", pub_cadena)
  PS_REP01.rdoParameters(0) = " "
  PS_REP01.rdoParameters(1) = 0
  PS_REP01.MaxRows = 1
  Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
  PS_REP01(0) = LK_CODCIA
  PS_REP01(1) = PUB_CODART
  llave_rep01.Requery
  If Not llave_rep01.EOF Then
     Screen.MousePointer = 0
     MsgBox "NO se Puede Eliminar ...  ARTICULO  TIENE H I S T O R I A.. ", 48, Pub_Titulo
     Exit Sub
  End If
  
  If LK_EMP_PTO = "A" Then
    If LK_CODCIA <> "00" Then
      Screen.MousePointer = 0
      LblMensaje.Visible = False
      MsgBox "No Procede la Eliminación.  Punto de Venta no permitido!!(solo en la Cia. central)", 48, Pub_Titulo
      Exit Sub
    End If
  End If
  LblMensaje.Visible = False
  pub_mensaje = " ¿Desea Eliminar el Articulo... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then
      LblMensaje.Visible = False
      Screen.MousePointer = 0
      Exit Sub
  End If
  LblMensaje.Visible = True
  LblMensaje.Caption = "Eliminando.  un Momento..."
  
  CN.Execute "Begin Transaction", rdExecDirect
  pub_cadena = "SELECT * FROM CONTROLL"
  Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
  If LK_EMP_PTO = "A" Then
     If Trim(GEN!gen_ART_CIAS) <> "" Then
        xcuenta = 1
        For fila = 1 To 30
            ws_codcia = Mid(Trim(GEN!gen_ART_CIAS), xcuenta, 2)
            If ws_codcia = "00" Then
               flag_puntos = "A"
            Else
               flag_puntos = ""
            End If
            If Trim(ws_codcia) = "" Then Exit For
            GoSub BORRA_ART_CIA
            xcuenta = xcuenta + 2
        Next fila
     Else
        MsgBox "Verificar.!!! NO esta activado la opcion de Puntos de Venta (está Cia esta Activada), Consultar al Administrador.", 48, Pub_Titulo
        LblMensaje.Visible = False
        Screen.MousePointer = 0
        Exit Sub
     End If
  Else
    ws_codcia = LK_CODCIA
    flag_puntos = "A"
    GoSub BORRA_ART_CIA
  End If

CN.Execute "Commit Transaction", rdExecDirect
con_llave.Close

cmdcancelar_Click
MENSAJE_ARTI "Articulo   ELIMINADO ... "
Screen.MousePointer = 0
Exit Sub

BORRA_ART_CIA:
If flag_puntos = "A" Then
 SQ_OPER = 1
 PUB_KEY = WS_CODART
 pu_codcia = ws_codcia
 LEER_ART_LLAVE
 art_LLAVE.Delete
End If
 SQ_OPER = 1
 PUB_CODART = WS_CODART
 pu_codcia = ws_codcia
 LEER_ARM_LLAVE
 arm_llave.Delete
 pub_cadena = "DELETE PRECIOS WHERE PRE_CODCIA = '" & ws_codcia & "' AND PRE_CODART = " & WS_CODART
 CN.Execute pub_cadena, rdExecDirect
Return

Exit Sub

ESCAPA:
    con_llave.Close
    CN.Execute "Rollback Transaction", rdExecDirect
    Screen.MousePointer = 0
    MsgBox Err.Number & "  " & Err.Description & "  Intente Nuevamente."
    LblMensaje.Visible = False
    DoEvents
    cmdCancelar.Enabled = True
    cmdEliminar.Enabled = True
    cmdModificar.Enabled = True
    cmdAgregar.Enabled = True
    LIMPIA_ARTI
    BLOQUEA_TEXT art_linea, art_numero, art_marca, art_plancha, checambio, txtlitro
    BLOQUEA_TEXT txtnombre, CmbCalidad, decimales, DS, txtcospro, art_situacion, exigv, txtcospro, cmddolares, txtfechault
    BLOQUEA_TEXT art_grupo, art_familia, art_subfam, grid_unid, txtMin, txtMax, art_codpro, txtcodigo2, cmddolares, txtpeso
    BLOQUEA_TEXT txtcolor, txtmarca
    If LK_EMP = "HER" Then
           BLOQUEA_TEXT txtpor1, txtpor2, txtpor3, txtpor4, txtpor5, txtpor6
    End If
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
     DESBLOQUEA_TEXT txt_alterno
     txt_alterno.SetFocus
    Else
     DESBLOQUEA_TEXT Txt_key
     Txt_key.SetFocus
    End If
    Screen.MousePointer = 0

End Sub

Private Sub cmdEliminar_GotFocus()
If ListView1.Visible Then
  ListView1.Visible = False
  Txt_key.Text = ""
End If
End Sub

Private Sub cmdEliminar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
'    frmARTI.txt_key.SetFocus
End If

End Sub

Private Sub CmdEscapa_Click()
If (frmARTI.txtnombre.Enabled = True) Then
  frmARTI.txtnombre.SetFocus
  End If
  frmARTI.F14.Visible = False
  updatePrecios = False
End Sub

Private Sub CmdModificar_Click()
'On Error GoTo ESCAPA
If Trim(Txt_key.Text) = "1" Then
     MENSAJE_ARTI "No Procede. .."
     Exit Sub
End If
If Len(Txt_key) = 0 Or Trim(txtnombre.Text) = "" Then
   MENSAJE_ARTI "NO a seleccionado ningun Articulo... !"
   Exit Sub
End If
If Left(cmdModificar.Caption, 2) = "&M" Then
    cmdModificar.Caption = "&Grabar"
    cmdEliminar.Enabled = False
    cmdAgregar.Enabled = False
    cmdCancelar.Enabled = True
    DESBLOQUEA_TEXT txt_alterno
    'If LK_CODUSU = "SUPERVISOR" Or LK_CODUSU = "ADMIN" Then
    '   DESBLOQUEA_TEXT txt_alterno
    'Else
    '   BLOQUEA_TEXT txt_alterno
    'End If
    BLOQUEA_TEXT Txt_key
    DESBLOQUEA_TEXT txtnombre, decimales, DS, txtcospro, art_situacion, art_linea, art_numero, art_marca, art_plancha
    DESBLOQUEA_TEXT art_grupo, art_familia, art_subfam, grid_unid, txtMin, txtMax, art_codpro, txtcodigo2
    DESBLOQUEA_TEXT cheservi(0), cheservi(1), cheservi(2), exigv, txtcospro, cmddolares, txtpeso, txtfechault, checambio, txtlitro
    DESBLOQUEA_TEXT txtcolor, txtmarca
    If LK_EMP_PTO = "A" Then
      If LK_CODCIA <> "00" Then
        BLOQUEA_TEXT decimales, art_grupo, art_familia, art_subfam, art_codpro, txtcodigo2
        BLOQUEA_TEXT art_situacion, DS, cheservi(0), cheservi(1), cheservi(2), txtMin, txtMax
      End If
    End If
    If LK_EMP = "HER" Then
       DESBLOQUEA_TEXT txtpor1, txtpor2, txtpor3, txtpor4, txtpor5, txtpor6
    End If
    'MODIFICADO PARA PODER MODIFICAR LA MONEDA
    MANOS(0).Enabled = False
    MANOS(1).Enabled = False
    '07/12/2001
    If LK_MONEDA = "D" Then
        frmARTI.DS.Enabled = Not False
    ElseIf LK_MONEDA = "S" Then
        frmARTI.DS.Enabled = Not False
    End If
    If loc_flag_bloq = "A" Then
      fcomun.Enabled = False
    Else
      fcomun.Enabled = True
    End If
    pasa = 1
    updatePrecios = False
    frmARTI.txtnombre.SetFocus
Else
    '*Grabar las modificaciones
    If Not IsDate(frmARTI.txtfechault.Text) Then
       MsgBox "Fecha de ultima compra no es correcta.", 48, Pub_Titulo
      Exit Sub
    End If
    Screen.MousePointer = 11
    If Trim(DS.Text) = "D" Or Trim(DS.Text) = "S" Then
    Else
      Screen.MousePointer = 0
      MsgBox "Solo Acepta  D = Dolares,  o  S = Soles ... ", 48, Pub_Titulo
      DS.SetFocus
      Exit Sub
    End If
    If Val(decimales.Text) < 0 Then
      Screen.MousePointer = 0
      MsgBox "Escoger con cuantos decimales de precision...", 48, Pub_Titulo
      decimales.SetFocus
      Exit Sub
    End If
    If Not CONSIS_ARTI Then
       Screen.MousePointer = 0
       Exit Sub
    End If
    If Not CONSIS_UNIDAD Then
       Screen.MousePointer = 0
       grid_unid.SetFocus
       Exit Sub
    End If
    If pasa = 1 Then
      If EXISTE_ART(txtnombre.Text, Trim(Txt_key.Text)) Then
          Screen.MousePointer = 0
          MENSAJE_ARTI "Existen algunos Articulos con estos NOMBRES .."
          frmARTI.ListExiste.SetFocus
          Exit Sub
      End If
    End If
    pasa = 0
    CN.Execute "Begin Transaction", rdExecDirect
    If Trim(Nulo_Valors(par_llave!par_art_cias)) <> "" Then
        xcuenta = 1
        For fila = 1 To 30
          pu_codcia = Mid(Trim(par_llave!par_art_cias), xcuenta, 2)
          If Trim(pu_codcia) = "" Then Exit For
             PUB_KEY = Val(frmARTI.Txt_key.Text)
             PUB_CODCIA = pu_codcia
             LOC_OPER = 1
             LEER_LOC
             GRABAR_ARTI
            'GoSub IR_POR_CIA
          xcuenta = xcuenta + 2
        Next fila
    Else
      GRABAR_ARTI
    End If
    CN.Execute "Commit Transaction", rdExecDirect
    If (updatePrecios = False) Then
     cmdModificar.Caption = "&Modificación"
     cmdCancelar.Enabled = True
     cmdEliminar.Enabled = True
     cmdAgregar.Enabled = True
    End If
    LIMPIA_ARTI
    BLOQUEA_TEXT txtnombre, CmbCalidad, decimales, DS, txtcospro, art_situacion, art_linea, art_numero, art_marca, art_plancha
    BLOQUEA_TEXT art_grupo, art_familia, art_subfam, grid_unid, txtMin, txtMax, art_codpro, txtcodigo2
    BLOQUEA_TEXT cheservi(0), cheservi(1), cheservi(2), exigv, txtcospro, cmddolares, txtpeso, txtfechault, checambio, txtlitro
    BLOQUEA_TEXT txtcolor, txtmarca
    If LK_EMP = "HER" Then
       BLOQUEA_TEXT txtpor1, txtpor2, txtpor3, txtpor4, txtpor5, txtpor6
    End If
    MENSAJE_ARTI "Articulo,  MODIFICADO... "
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      DESBLOQUEA_TEXT txt_alterno
      txt_alterno.SetFocus
    Else
      DESBLOQUEA_TEXT Txt_key
      Txt_key.SetFocus
    End If
    MANOS(0).Enabled = True
    MANOS(1).Enabled = True
    Screen.MousePointer = 0
End If
Exit Sub
ESCAPA:
    If Err.Number = 40002 Then
        MsgBox "Hay Error en la LLave ..Intente Nuevamente. "
    Else
        MsgBox Err.Number & "  " & Err.Description & "  Intente Nuevamente."
    End If
    CN.Execute "Rollback Transaction", rdExecDirect
    cmdModificar.Caption = "&Modificación"
    cmdCancelar.Enabled = True
    cmdEliminar.Enabled = True
    cmdModificar.Enabled = True
    cmdAgregar.Enabled = True
    LIMPIA_ARTI
    BLOQUEA_TEXT txtnombre, CmbCalidad, decimales, DS, txtcospro, art_situacion, art_linea, art_numero, art_marca, art_plancha
    BLOQUEA_TEXT art_grupo, art_familia, art_subfam, grid_unid, txtMin, txtMax, art_codpro, txtcodigo2
    BLOQUEA_TEXT cheservi(0), cheservi(1), cheservi(2), exigv, txtcospro, cmddolares, txtpeso, txtfechault, checambio, txtlitro
    BLOQUEA_TEXT txtcolor, txtmarca
    If LK_EMP = "HER" Then
       BLOQUEA_TEXT txtpor1, txtpor2, txtpor3, txtpor4, txtpor5, txtpor6
    End If
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
     DESBLOQUEA_TEXT txt_alterno
     txt_alterno.SetFocus
    Else
     DESBLOQUEA_TEXT Txt_key
     Txt_key.SetFocus
    End If
    Screen.MousePointer = 0
    MANOS(0).Enabled = True
    MANOS(1).Enabled = True

End Sub

Private Sub cmdModificar_GotFocus()
If ListView1.Visible Then
  ListView1.Visible = False
  Txt_key.Text = ""
End If
End Sub

Private Sub cmdModificar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 '   frmARTI.txt_key.SetFocus
End If
End Sub


Private Sub cmdp_Click()
Dim wnombre As String
Dim WORIGINAL As Currency
Dim WORIGINAL1 As Currency
Dim WORIGINAL2 As Currency
Dim walterno  As String
Dim WCALIDAD  As Integer
Dim WFAMILIA  As Integer
Dim WSUBFAMI  As Integer
Dim WNUMERO As Integer
Dim wgrupo As Integer
Dim WLINEA As Integer
Dim wMARCA As Integer

If Trim(cmbcal.Text) = "" Then
 MsgBox " Seleccione su Calidad.", 48, Pub_Titulo
 cmbcal.SetFocus
 Exit Sub
End If
If Val(Right(art_familia.Text, 8)) <> 4 Then
  MsgBox "solo se Activa para los 4 - Producto Terminado", 48, Pub_Titulo
  Exit Sub
End If

For fila = 2 To 3
    WORIGINAL = artloc_llave!ART_KEY
    If fila = 2 Then
      WORIGINAL1 = GENERA_CODI()
    Else
      WORIGINAL2 = GENERA_CODI()
    End If
'    walterno = WORIGINAL
    WCALIDAD = 1
    If fila = 2 Then
      WFAMILIA = 3
    Else
      WFAMILIA = 2
    End If
    WSUBFAMI = 1
    wnombre = artloc_llave!art_nombre
    For i = 0 To art_familia.ListCount - 1
      art_familia.ListIndex = i
      If WFAMILIA = Val(Right(art_familia.Text, 6)) Then
         wnombre = Trim(Left(art_familia.Text, 6)) & " " & Trim(Left(art_grupo.Text, 15)) & " " & Trim(Left(art_numero.Text, 15)) & "-" & Trim(Left(art_marca.Text, 15))  ' & " " & Left(art_linea.Text, 3) ' Trim(Left(art_linea.Text, 5))
         Exit For
      End If
    Next i
    WNUMERO = artloc_llave!art_numero
    wgrupo = artloc_llave!art_subgru
    WLINEA = artloc_llave!art_linea
    wMARCA = artloc_llave!art_marca
    GoSub grabar
Next fila
PSART_RELA.rdoParameters(0) = LK_CODCIA
PSART_RELA.rdoParameters(1) = WORIGINAL1
art_rela.Requery
If art_rela.EOF Then
  MsgBox "NO ESTA BIEN LA RELACION."
  Exit Sub
End If
art_rela.Edit
art_rela!ART_CODART2 = WORIGINAL2
art_rela.Update

artloc_llave.Edit
artloc_llave!ART_CODART2 = WORIGINAL1
artloc_llave.Update
cmdcancelar_Click
MsgBox "Se Activo.", 48, Pub_Titulo
Exit Sub


grabar:

art_LLAVE.AddNew
If fila = 2 Then
    art_LLAVE!ART_KEY = WORIGINAL1
    art_LLAVE!art_alterno = Trim(Str(WORIGINAL1))
    art_LLAVE!ART_CODART2 = 0
Else
    art_LLAVE!ART_KEY = WORIGINAL2
    art_LLAVE!art_alterno = Trim(Str(WORIGINAL2))
    art_LLAVE!ART_CODART2 = 0
End If

art_LLAVE!art_familia = WFAMILIA
art_LLAVE!art_subfam = WSUBFAMI
art_LLAVE!ART_CALIDAD = WCALIDAD
art_LLAVE!art_subgru = wgrupo
art_LLAVE!art_linea = WLINEA
art_LLAVE!art_numero = WNUMERO
art_LLAVE!art_marca = wMARCA

art_LLAVE!art_nombre = wnombre
    
art_LLAVE!ART_POR_IGV = artloc_llave!ART_POR_IGV
art_LLAVE!ART_ORDEN = artloc_llave!ART_ORDEN
art_LLAVE!art_tipo = loc_tipo
art_LLAVE!art_codclie = artloc_llave!art_codclie
art_LLAVE!ART_CODCIA = PUB_CODCIA
art_LLAVE!ART_DECIMALES = artloc_llave!ART_DECIMALES
art_LLAVE!ART_MONEDA = artloc_llave!ART_MONEDA
art_LLAVE!art_situacion = artloc_llave!art_situacion
art_LLAVE!ART_STOCK_MIN = artloc_llave!ART_STOCK_MIN
art_LLAVE!ART_STOCK_MAX = artloc_llave!ART_STOCK_MAX
art_LLAVE!ART_POR_IGV = artloc_llave!ART_POR_IGV
art_LLAVE!ART_EX_IGV = ""
art_LLAVE!ART_COSPRO = 0
art_LLAVE!ART_EX_IGV = "A"
art_LLAVE!art_flag_stock = ""
art_LLAVE!art_flag_stock = "M"
art_LLAVE!ART_POR1 = 0
art_LLAVE!ART_POR2 = 0
art_LLAVE!ART_POR3 = 0
art_LLAVE!ART_POR4 = 0
art_LLAVE!ART_POR5 = 0
art_LLAVE.Update

'pu_codcia = PUB_CODCIA
'PUB_CODART = WORIGINAL
'SQ_OPER = 1
'LEER_ARM_LLAVE
'If arm_llave.EOF Then
    arm_llave.AddNew
If fila = 2 Then
    arm_llave!ARM_CODART = WORIGINAL1
Else
    arm_llave!ARM_CODART = WORIGINAL2
End If



    arm_llave!ARM_CODCIA = LK_CODCIA
    arm_llave!ARM_STOCK = 0
    arm_llave!ARM_INGRESOS = 0
    arm_llave!ARM_SALIDAS = 0
    arm_llave!arm_cospro = 0
    arm_llave!arm_stock2 = 0
    arm_llave!ARM_COSTO_ULT = 0
    arm_llave!ARM_FECHA_ULT = LK_FECHA_DIA
    arm_llave.Update
'End If

pre_mayor.AddNew
pre_mayor!PRE_CODCIA = LK_CODCIA
If fila = 2 Then
pre_mayor!PRE_codart = WORIGINAL1
Else
pre_mayor!PRE_codart = WORIGINAL2
End If

pre_mayor!pre_secuencia = 0
pre_mayor!pre_unidad = Left(grid_unid.TextMatrix(1, 0), 15)
pre_mayor!PRE_EQUIV = Val(grid_unid.TextMatrix(1, 1))
pre_mayor!pre_pre11 = Val(grid_unid.TextMatrix(1, 16))
pre_mayor!PRE_PRE22 = Val(grid_unid.TextMatrix(1, 17))
pre_mayor!PRE_PRE33 = Val(grid_unid.TextMatrix(1, 18))
pre_mayor!PRE_PRE44 = Val(grid_unid.TextMatrix(1, 19))
pre_mayor!PRE_PRE55 = Val(grid_unid.TextMatrix(1, 20))
''**********************************PRECIO6
pre_mayor!PRE_PRE66 = Val(grid_unid.TextMatrix(1, 31))
''**********************************
pre_mayor!PRE_PRE1 = Val(grid_unid.TextMatrix(1, 21))
pre_mayor!PRE_PRE2 = Val(grid_unid.TextMatrix(1, 22))
pre_mayor!PRE_PRE3 = Val(grid_unid.TextMatrix(1, 23))
pre_mayor!PRE_PRE4 = Val(grid_unid.TextMatrix(1, 24))
pre_mayor!PRE_PRE5 = Val(grid_unid.TextMatrix(1, 25))
''**********************************PRECIO6
pre_mayor!PRE_PRE6 = Val(grid_unid.TextMatrix(1, 32))
''**********************************
pre_mayor!pre_PESO = Val(grid_unid.TextMatrix(1, 26))
pre_mayor!PRE_LITRO = Val(grid_unid.TextMatrix(1, 28))
pre_mayor!PRE_FLAG_UNIDAD = grid_unid.TextMatrix(1, 14)
pre_mayor.Update
Return
End Sub

Private Sub cmdpath_Click()
With cdlfoto
On Error GoTo ErrorHandle
  .FLAGS = cdlOFNHideReadOnly
  .Filter = "Archivos de Imagenes |*.jpg|*.bmp"
  .FilterIndex = 2
  .ShowOpen
  txtpath = .FileName
  picfoto.Picture = LoadPicture(.FileName)
End With
Exit Sub
ErrorHandle:
 picfoto.Picture = LoadPicture()
 txtpath = ""
End Sub

Private Sub cmdquitar_Click()
On Error GoTo ESCAPA
SQ_OPER = 1
PUB_CODART = Val(gridrel.TextMatrix(1, 0))
PUB_KEY = PUB_CODART
pu_codcia = LK_CODCIA
LEER_ART_LLAVE
If art_LLAVE.EOF Then
  MsgBox "Arti no Existe ,,", 48, Pub_Titulo
  Exit Sub
End If

SQ_OPER = 1
PUB_CODART = Val(gridrel.TextMatrix(1, 0))
pu_codcia = LK_CODCIA
LEER_ARM_LLAVE
If arm_llave.EOF Then
  MsgBox "Arti no Existe ,,", 48, Pub_Titulo
  Exit Sub
End If
If arm_llave!ARM_INGRESOS = 0 And arm_llave!ARM_SALIDAS = 0 Then
Else
  MsgBox "NO Procede - Articulo Defectuoso Tiene Historia ", 48, Pub_Titulo
  Exit Sub
End If
If art_LLAVE!ART_CALIDAD <> 2 Then
  MsgBox "Articulo no Es Defectuoso Verificar !!!!! ", 48, Pub_Titulo
  Exit Sub
End If

pub_cadena = "SELECT FAR_CODART FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODART = ?  AND FAR_ESTADO <> 'E' "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01.rdoParameters(0) = " "
PS_REP01.rdoParameters(1) = 0
PS_REP01.MaxRows = 1
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = PUB_CODART
llave_rep01.Requery
If Not llave_rep01.EOF Then
   Screen.MousePointer = 0
   MsgBox "NO se Puede Eliminar ...  ARTICULO  TIENE H I S T O R I A.. ", 48, Pub_Titulo
   Exit Sub
End If
  

pub_mensaje = " Desea Eliminar la Relacion y el Articulo. :" & Trim(gridrel.TextMatrix(1, 2)) & " ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
  Exit Sub
End If


CN.Execute "Begin Transaction", rdExecDirect
PUB_CODART = Val(gridrel.TextMatrix(1, 0))
CN.Execute "DELETE ARTI WHERE ART_KEY = " & PUB_CODART & " AND ART_CODCIA = '" & LK_CODCIA & "'", rdExecDirect
CN.Execute "DELETE PRECIOS WHERE PRE_CODART = " & PUB_CODART & " AND PRE_CODCIA = '" & LK_CODCIA & "'", rdExecDirect
CN.Execute "DELETE ARTICULO WHERE ARM_CODART = " & PUB_CODART & " AND ARM_CODCIA = '" & LK_CODCIA & "'", rdExecDirect
CN.Execute "Commit Transaction", rdExecDirect
On Error GoTo 0
gridrel.TextMatrix(1, 0) = ""
gridrel.TextMatrix(1, 1) = ""
gridrel.TextMatrix(1, 2) = ""
gridrel.TextMatrix(1, 3) = ""
cmdquitar.Visible = False
cmdAdd.Visible = True
cmbcal.Visible = True
lblcal.Visible = True
Exit Sub
ESCAPA:
CN.Execute "Rollback Transaction", rdExecDirect
MsgBox "Intente Nuevamente.", 48, titulo

End Sub

Private Sub decimales_GotFocus()
frmARTI.F14.Visible = False
End Sub

Private Sub decimales_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Flag_Inicial = "A"
 grid_unid.COL = 0
 grid_unid.Row = 1
 Flag_Inicial = ""
 grid_unid.SetFocus
End If

End Sub

Private Sub decimales_LostFocus()
Fvarios.Refresh
End Sub

Private Sub DS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Flag_Inicial = "A"
 grid_unid.COL = 0
 grid_unid.Row = 1
 Flag_Inicial = ""
 If Trim(frmARTI.DS.Text) = "S" Then
     cmddolares.Caption = "Lista de Precios en S/. (Nuevos Soles)"
     cmddolares.Tag = "S"
 Else
     cmddolares.Caption = "Lista de Precios en US$. (Dolares Americanos)"
     cmddolares.Tag = "D"
 End If
 'cmddolares_Click
 If grid_unid.Enabled Then grid_unid.SetFocus
  
End If
End Sub

Private Sub exigv_Click()
If Left(cmdModificar.Caption, 2) = "&M" And Left(cmdAgregar.Caption, 2) = "&A" Then Exit Sub
  
If exigv.Value = 0 Then
'  txtcospro.Enabled = False
  txtcospro.Text = ""
Else
'  txtcospro.Enabled = True
  txtcospro.SetFocus
End If
End Sub


Private Sub Form_Activate()
frmARTI.SSTab1.tab = 0
End Sub

Private Sub Form_DblClick()
opcional
End Sub

Private Sub Form_Load()
'para o/C agrgado por mic
    SETGRID
    'para buscador
    LlenadoCbo artfamilia, 122
    LlenadoCbo artgrupo, 129
    LlenadoCbo artlinea, 131
    artlinea.AddItem " < -- TODAS -->                                                    -1"
'***********
For fila = 1 To lk_OTROS_Count
   If Val(lk_OTROS(fila)) = 6 Then ' bloque de precios en mastros de articulos
    loc_flag_bloq = "A"
   End If
Next fila

 
pub_cadena = "SELECT ART_NOMBRE, ART_ALTERNO, ART_KEY , ART_codart2  FROM ARTI WHERE ART_CODCIA = ? AND ART_KEY = ? "
Set PSART_RELA = CN.CreateQuery("", pub_cadena)
PSART_RELA.rdoParameters(0) = 0
PSART_RELA.rdoParameters(1) = 0
Set art_rela = PSART_RELA.OpenResultset(rdOpenKeyset, rdConcurValues)


pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ? AND ART_KEY = ? ORDER BY ART_KEY"
Set PSART_KEY = CN.CreateQuery("", pub_cadena)
PSART_KEY.rdoParameters(0) = "  "
PSART_KEY.rdoParameters(1) = 0
Set artloc_key = PSART_KEY.OpenResultset(rdOpenKeyset, rdConcurValues)




If LK_EMP = "HER" Then
 picfoto.Visible = True
  End If
  'Else
'  Fop.Left = 120
'End If
    VAR_NEWCAL = 0
    frmARTI.F14.Left = 90
    frmARTI.F14.Top = 3360
    frmARTI.F14.Height = 2415
    frmARTI.F14.Width = 9380
    pasa = 0
    frmARTI.ListExiste.Cols = 5
    For fila = 0 To frmARTI.ListExiste.COL - 1
      frmARTI.ListExiste.COL = fila
      frmARTI.ListExiste.FixedAlignment(fila) = 2
    Next fila
    pasa = 0
    frmARTI.ListExiste.ColWidth(0) = 350
    frmARTI.ListExiste.ColWidth(1) = 800
    frmARTI.ListExiste.ColWidth(2) = 4000
    frmARTI.ListExiste.ColWidth(3) = 1500
    Screen.MousePointer = 11
    frmARTI.DS.AddItem "S"
    frmARTI.DS.AddItem "D"
    frmARTI.decimales.AddItem "1"
    frmARTI.decimales.AddItem "2"
    frmARTI.decimales.AddItem "3"
    frmARTI.decimales.AddItem "4"
    frmARTI.DS.ListIndex = -1
    frmARTI.decimales.ListIndex = -1
    BLOQUEA_TEXT txtnombre, CmbCalidad, decimales, DS, txtcospro, art_situacion, art_linea, art_numero, art_marca, art_plancha
    BLOQUEA_TEXT art_grupo, art_familia, art_subfam, grid_unid, txtMin, txtMax, art_codpro, txtcodigo2
    BLOQUEA_TEXT cheservi(0), cheservi(1), cheservi(2), exigv, txtcospro, cmddolares, txtpeso, txtfechault, checambio, txtlitro
    BLOQUEA_TEXT txtcolor, txtmarca
    If LK_EMP = "HER" Then
       BLOQUEA_TEXT txtpor1, txtpor2, txtpor3, txtpor4, txtpor5, txtpor6
    End If
    If LK_FLAG_ORIGINAL = "A" Then
       txt_alterno.Visible = False
       lblalterno.Visible = False
       txtnombre.Left = 2160
       lblnomarti.Left = 2160
    End If
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       BLOQUEA_TEXT Txt_key
       pub_cadena = "SELECT ART_ALTERNO FROM ARTI WHERE ART_CODCIA = ? AND ART_KEY <> 0 order by art_alterno"
    Else
      BLOQUEA_TEXT txt_alterno
       pub_cadena = "SELECT ART_KEY FROM ARTI WHERE ART_CODCIA = ? AND ART_KEY <> 0 order by art_NOMBRE"
    End If
    Set PSMANO_CODI = CN.CreateQuery("", pub_cadena)
    PSMANO_CODI.rdoParameters(0) = 0
    Set mano_CODI = PSMANO_CODI.OpenResultset(rdOpenKeyset, rdConcurValues)
    PSMANO_CODI(0) = LK_CODCIA
    mano_CODI.Requery
    
    Fvarios.Visible = False
    fcomun.Visible = True
    Fdatos.Visible = True
    PROCESO_ARTI
    loc_tipo = "V"
    LLENADO_FAM
    'LLENADO_GRUPO
    LLENADO_CAL
    'LLENADO_NUMERO
    LLENADO_MARCA
    LLENADO_PLANCHA
    LLENADO_LINEA 131
    
    Fvarios.Visible = True
    Screen.MousePointer = 0
    lblUnidad.Caption = ""
    frmARTI.fcomun.Visible = True
    frmARTI.fcomun.Enabled = True
    grid_unid.Enabled = False
    cmdCancelar.Enabled = True
    SQ_OPER = 2
    PUB_TIPREG = 45
    PUB_CODCIA = LK_CODCIA
    LEER_TAB_LLAVE
    Do Until tab_mayor.EOF
      Label3(tab_mayor!TAB_NUMTAB - 1).Caption = Trim(tab_mayor!tab_NOMLARGO)
      lblpor(tab_mayor!TAB_NUMTAB - 1).Caption = Left(lblpor(tab_mayor!TAB_NUMTAB - 1).Caption, 5) & Trim(tab_mayor!tab_NOMLARGO) & " :"
      tab_mayor.MoveNext
   Loop
   pasa = 0
   loc_tipo = "V"
   PROCESA_PROV
   frmARTI.frarelacion.Enabled = False
   cmddolares.Caption = "Lista de Precios en S/. (Nuevos Soles)"
   cmddolares.Tag = "S"
   SQ_OPER = 1
   PUB_CODCIA = "00"
   PUB_TIPREG = 340
   PUB_NUMTAB = 0
   LEER_TAB_LLAVE
   If Not tab_llave.EOF Then lblart(0).Caption = Trim(tab_llave!tab_NOMLARGO)
   PUB_NUMTAB = 1
   LEER_TAB_LLAVE
   If Not tab_llave.EOF Then lblart(1).Caption = Trim(tab_llave!tab_NOMLARGO)
   PUB_NUMTAB = 2
   LEER_TAB_LLAVE
   If Not tab_llave.EOF Then lblart(2).Caption = Trim(tab_llave!tab_NOMLARGO)
   PUB_NUMTAB = 3
   LEER_TAB_LLAVE
   If Not tab_llave.EOF Then lblart(3).Caption = Trim(tab_llave!tab_NOMLARGO)
   PUB_NUMTAB = 4
   LEER_TAB_LLAVE
   If Not tab_llave.EOF Then lblart(4).Caption = Trim(tab_llave!tab_NOMLARGO)
   PUB_NUMTAB = 5
   LEER_TAB_LLAVE
   If Not tab_llave.EOF Then lblart(5).Caption = Trim(tab_llave!tab_NOMLARGO)
   'PUB_NUMTAB = 6
   'LEER_TAB_LLAVE
   'If Not tab_llave.EOF Then lblart(5).Caption = Trim(tab_llave!TAB_NOMLARGO)
    If Trim(LK_CODUSU) = "ADMIN" Then
        cmd_AddItem.Visible = True
    End If
End Sub

Public Sub LLENADO_FAM()
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = 122
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       PUB_CODCIA = "00"
    End If
    SQ_OPER = 2
    LEER_TAB_LLAVE
    art_familia.ToolTipText = "TAB_TIPREG = 122"
    art_familia.Clear
    Do Until tab_mayor.EOF
        art_familia.AddItem tab_mayor!tab_NOMLARGO & String(80, " ") & Trim(CStr(tab_mayor!TAB_NUMTAB))
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub

Public Sub LLENADO_SUBFAM(ctlCombo As ComboBox, ByVal wfami As Integer)
On Error GoTo SALE
Dim CONTA As Integer
    CONTA = -1
    Select Case ctlCombo.Name
      Case Is = "art_subfam"
       PUB_TIPREG = 123
'      Case Is = "art_grupo"
'       PUB_TIPREG = 129
'      Case Is = "art_numero"
'       PUB_TIPREG = 130
'      Case Is = "art_linea"
'       PUB_TIPREG = 131
    End Select
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
    'ctlCombo.ToolTipText = "TAB_TIPREG = 123"
    ctlCombo.Clear
    Do Until tab_menor.EOF
        ctlCombo.AddItem tab_menor!tab_NOMLARGO & String(80, " ") & Trim(CStr(tab_menor!TAB_NUMTAB))
        DoEvents
        CONTA = CONTA + 1
        tab_menor.MoveNext
    Loop
Exit Sub
SALE:
Resume Next
End Sub
Public Sub LLENADO_GRUPO(ByVal wSubFam As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = 129
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       PUB_CODCIA = "00"
    End If
    SQ_OPER = 2
    PUB_CODART = wSubFam
    LEER_TAB_LLAVE
    art_grupo.ToolTipText = "TAB_TIPREG = 129"
    art_grupo.Clear
    Do Until tab_mayor.EOF
        art_grupo.AddItem tab_mayor!tab_NOMLARGO & String(80, " ") & Trim(CStr(tab_mayor!TAB_NUMTAB))
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub
'PROCEDIMIENTO PARA LLENAR GRUPO 2
Public Sub LLENADO_NUMERO(ByVal wgrupo As Integer)
    PUB_TIPREG = 130
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       PUB_CODCIA = "00"
    End If
    SQ_OPER = 2
    PUB_CODART = wgrupo
    LEER_TAB_LLAVE
    art_numero.ToolTipText = "TAB_TIPREG = 130"
    art_numero.Clear
    Do Until tab_mayor.EOF
       art_numero.AddItem tab_mayor!tab_NOMLARGO & String(80, " ") & Trim(CStr(tab_mayor!TAB_NUMTAB))
       tab_mayor.MoveNext
    Loop
End Sub
Public Sub LLENADO_MARCA()
    PUB_TIPREG = 132
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       PUB_CODCIA = "00"
    End If
    SQ_OPER = 2
    LEER_TAB_LLAVE
    art_marca.ToolTipText = "TAB_TIPREG = 132"
    art_marca.Clear
    Do Until tab_mayor.EOF
       art_marca.AddItem tab_mayor!tab_NOMLARGO & String(80, " ") & Trim(CStr(tab_mayor!TAB_NUMTAB))
       tab_mayor.MoveNext
    Loop
End Sub
Public Sub LLENADO_PLANCHA()
    PUB_TIPREG = 133
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       PUB_CODCIA = "00"
    End If
    SQ_OPER = 2
    LEER_TAB_LLAVE
    art_plancha.ToolTipText = "TAB_TIPREG = 133"
    art_plancha.Clear
    Do Until tab_mayor.EOF
       art_plancha.AddItem tab_mayor!tab_NOMLARGO & String(80, " ") & Trim(CStr(tab_mayor!TAB_NUMTAB))
       tab_mayor.MoveNext
    Loop
End Sub
'PROCEDIMEINTO PARA LLENAR LAS LINEAS
Public Sub LLENADO_LINEA(ByVal wSubGrupo As Integer)
    PUB_TIPREG = 131
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       PUB_CODCIA = "00"
    End If
    SQ_OPER = 2
    PUB_CODART = wSubGrupo
    LEER_TAB_LLAVE
    art_linea.ToolTipText = "TAB_TIPREG = 131"
    art_linea.Clear
    Do Until tab_mayor.EOF
       art_linea.AddItem tab_mayor!tab_NOMLARGO & String(80, " ") & Trim(CStr(tab_mayor!TAB_NUMTAB))
       tab_mayor.MoveNext
    Loop
End Sub
Public Sub LLENADO_CAL()
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = 2
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       PUB_CODCIA = "00"
    End If
    SQ_OPER = 2
    LEER_TAB_LLAVE
    CmbCalidad.ToolTipText = "TAB_TIPREG = 2"
    CmbCalidad.Clear
    Do Until tab_mayor.EOF
        CmbCalidad.AddItem tab_mayor!tab_NOMLARGO & String(80, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub
Public Sub ASIGNA(WCONTROL As ComboBox, txt As String)
Dim C As Integer
For C = 0 To WCONTROL.ListCount - 1
    If Trim(WCONTROL.List(C)) = Trim(txt) Then
        WCONTROL.ListIndex = C
        Exit Sub
    End If
Next C
End Sub
Public Sub LEER_LOC()
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If
If LOC_OPER = 1 Then
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      PSART_LOC.rdoParameters(0) = pu_alterno
    Else
      PSART_LOC.rdoParameters(0) = PUB_KEY
    End If
    PSART_LOC.rdoParameters(1) = PUB_CODCIA
    PSART_LOC.rdoParameters(2) = loc_tipo
    artloc_llave.Requery
ElseIf LOC_OPER = 2 Then
  PSART_KEY.rdoParameters(0) = PUB_CODCIA
  PSART_KEY.rdoParameters(1) = PUB_KEY
  artloc_key.Requery
End If

End Sub

Public Sub CAB_ARTI()
grid1.TextMatrix(0, 0) = "Cia."
grid1.TextMatrix(0, 3) = "Grupo"
grid1.TextMatrix(0, 2) = "Articulo"

End Sub
Public Sub ASIGNA_CHAR(WCONTROL As ComboBox, txt As String)
Dim C As Integer
For C = 0 To WCONTROL.ListCount - 1
    If Trim(Left(WCONTROL.List(C), 1)) = Trim(txt) Then
        WCONTROL.ListIndex = C
        Exit Sub
    End If
Next C
End Sub
Public Sub ASIGNA_INT(WCONTROL As ComboBox, txt As Integer)
Dim C As Integer
For C = 0 To WCONTROL.ListCount - 1
    If Val(Trim(Right(WCONTROL.List(C), 6))) = txt Then
        WCONTROL.ListIndex = C
        Exit Sub
    End If
Next C
End Sub




Private Sub grid_unid_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then
 Azul txtpeso, txtpeso
End If
End Sub

Private Sub Label3_DblClick(Index As Integer)
If Trim(LK_CODUSU) <> "ADMIN" And Trim(LK_CODUSU) <> "SUPERVISOR" Then
 Exit Sub
End If
If Trim(Label3(Index).Tag) = "" Then
 Exit Sub
End If
Dim wnombre
wnombre = InputBox("Ingrese la Nueva Descripción para este Campo :", Pub_Titulo, Trim(Label3(Index).Caption))
If wnombre = "" Then
  Screen.MousePointer = 0
  Exit Sub
End If
Screen.MousePointer = 11
SQ_OPER = 1
PUB_TIPREG = 45
PUB_NUMTAB = Val(Label3(Index).Tag)
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_llave.EOF Then
  tab_llave.AddNew
Else
  tab_llave.Edit
End If
  tab_llave!TAB_CODCIA = LK_CODCIA
  tab_llave!TAB_TIPREG = 45
  tab_llave!TAB_NUMTAB = Val(Label3(Index).Tag)
  tab_llave!tab_NOMLARGO = Left(wnombre, 40)
  tab_llave!tab_nomcorto = Left(wnombre, 10)
  tab_llave.Update
  Label3(Index).Caption = Left(wnombre, 40)
  lblpor(Index).Caption = Left(lblpor(Index).Caption, 5) & Trim(wnombre) & " :"
Screen.MousePointer = 0

End Sub

Private Sub ListExiste_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 And frmARTI.txtnombre.Enabled Then
    frmARTI.txtnombre.SetFocus
    frmARTI.F14.Visible = False
    Exit Sub
End If

End Sub

Private Sub ListExiste_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
    If ListExiste.TextMatrix(ListExiste.Row, 4) = "1" Then
     ListExiste.TextMatrix(ListExiste.Row, 4) = "0"
    Else
     ListExiste.TextMatrix(ListExiste.Row, 4) = "1"
    End If
End If

End Sub

Private Sub ListExiste_LostFocus()
If frmARTI.ListExiste.Visible = False Then
    Exit Sub
End If

End Sub

Private Sub ListView1_DblClick()
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
 loc_key = ListView1.SelectedItem.Index
 txt_alterno.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
 txt_alterno_KeyPress 13
Else
 loc_key = ListView1.SelectedItem.Index
 Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
 txt_key_KeyPress 13
End If
End Sub

Private Sub ListView1_GotFocus()
If loc_key <> 0 Then
 Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
 ListView1.ListItems.Item(loc_key).Selected = True
 ListView1.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = ListView1.SelectedItem.Index
 If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  txt_alterno.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
 Else
  Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
 End If
End If

End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 ListView1.Visible = False
 If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" And txt_alterno.Enabled Then
  txt_alterno.Text = ""
  txt_alterno.SetFocus
 ElseIf LK_FLAG_ALTERNO <> "A" And Txt_key.Enabled Then
  Txt_key.Text = ""
  Txt_key.SetFocus
 End If
 Exit Sub
End If
If KeyAscii <> 13 Then
 Exit Sub
End If
ListView1_DblClick
End Sub

Private Sub ListView1_LostFocus()
ListView1.Visible = False
End Sub


Private Sub MANOS_Click(Index As Integer)
mano_CODI.MoveFirst
If mano_CODI.RowCount > 0 Then
   Do Until mano_CODI.EOF
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       If Trim(mano_CODI!art_alterno) = Trim(txt_alterno.Text) Then Exit Do
    Else
       If Val(Txt_key.Text) = Val(mano_CODI!ART_KEY) Then Exit Do
    End If
    mano_CODI.MoveNext
  Loop
End If
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
    If txt_alterno.Text = "" Then
     mano_CODI.MoveFirst
    GoTo SALT
    End If
Else
    If Txt_key.Text = "" Then
     mano_CODI.MoveFirst
     GoTo SALT
    End If
End If


If Index = 0 Then
  If Not mano_CODI.BOF Then mano_CODI.MovePrevious
Else
  If Not mano_CODI.EOF Then mano_CODI.MoveNext
End If
SALT:
If mano_CODI.EOF Or mano_CODI.BOF Then Exit Sub
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  txt_alterno.Text = Trim(mano_CODI!art_alterno)
  txt_alterno_KeyPress 13
Else
 Txt_key.Text = Trim(mano_CODI!ART_KEY)
 txt_key_KeyPress 13

End If


End Sub

Private Sub PARPADEA_Timer()
 CU = CU + 1
 LblMensaje.Visible = Not LblMensaje.Visible
 If CU > 4 Then
   CU = 0
   PARPADEA.Enabled = False
   LblMensaje.Visible = False
 End If

End Sub

Public Sub LLENA_ARTI(ban As Integer)
Dim WFAMI2 As Integer
Dim WS_FLAG_UNIDAD As Integer
Dim WSPOR As Currency
If ban <> 1 Then
       If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
       Else
         If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
          If VAR_ACTIVAR = 99 Then
           txt_alterno.Text = Trim(ListView1.ListItems.Item(loc_key).Text)
          End If
          pu_alterno = Trim(txt_alterno.Text)
         Else
          If (updatePrecios = False) Then
           Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
           PUB_KEY = Val(Txt_key.Text)
          Else
           PUB_KEY = Val(Txt_key.Text)
          End If
         End If
       End If
       PUB_CODCIA = LK_CODCIA
       LOC_OPER = 1
       LEER_LOC
       If artloc_llave.EOF Then
        Exit Sub
       End If
End If
'LLENADO
frmARTI.frarelacion.Enabled = True
frmARTI.Txt_key.Text = artloc_llave!ART_KEY
frmARTI.txtnombre.Text = RTrim(Nulo_Valors(artloc_llave!art_nombre))
'frmARTI.picfoto.Container = frmARTI.txtpath.Container  gts
ASIGNA_INT art_familia, Nulo_Valor0(artloc_llave!art_familia)

PUB_TIPREG = 123
WFAMI2 = Val(Trim(Right(art_familia.Text, 6)))
LLENADO_SUBFAM art_subfam, WFAMI2
ASIGNA_INT art_subfam, Nulo_Valor0(artloc_llave!art_subfam)

PUB_TIPREG = 129
WFAMI2 = Val(Trim(Right(art_subfam.Text, 6)))
LLENADO_SUBFAM art_grupo, WFAMI2
ASIGNA_INT art_grupo, Nulo_Valor0(artloc_llave!art_subgru)

PUB_TIPREG = 130
WFAMI2 = Val(Trim(Right(art_grupo.Text, 6)))
LLENADO_SUBFAM art_numero, WFAMI2
ASIGNA_INT art_numero, Nulo_Valor0(artloc_llave!art_numero)

PUB_TIPREG = 131
'WFAMI2 = Val(Trim(Right(art_numero.Text, 6)))
LLENADO_LINEA 131
ASIGNA_INT art_linea, Nulo_Valor0(artloc_llave!art_linea)

PUB_TIPREG = 132
WFAMI2 = Val(Trim(Right(art_linea.Text, 6)))
LLENADO_SUBFAM art_marca, WFAMI2
ASIGNA_INT art_marca, Nulo_Valor0(artloc_llave!art_marca)

ASIGNA_INT CmbCalidad, Nulo_Valor0(artloc_llave!ART_CALIDAD)

ASIGNA_INT art_plancha, Nulo_Valor0(artloc_llave!art_plancha)
ASIGNA_INT art_codpro, Nulo_Valor0(artloc_llave!art_codclie)
frmARTI.txtcospro.Text = Nulo_Valor0(artloc_llave!ART_POR_IGV)
SQ_OPER = 1
PUB_CODART = artloc_llave!ART_KEY
pu_codcia = LK_CODCIA
LEER_ARM_LLAVE
frmARTI.lblcospro.Caption = Nulo_Valor0(arm_llave!arm_cospro)
frmARTI.txtfechault.Text = Format(arm_llave!ARM_FECHA_ULT, "dd/mm/yyyy")
frmARTI.DS.Text = Trim(Nulo_Valors(artloc_llave!ART_MONEDA))
frmARTI.decimales.Text = Val(Nulo_Valor0(artloc_llave!ART_DECIMALES))
frmARTI.txt_alterno.Text = Nulo_Valors(artloc_llave!art_alterno)
frmARTI.txtMin.Text = Nulo_Valors(artloc_llave!ART_STOCK_MIN)
frmARTI.txtMax.Text = Nulo_Valors(artloc_llave!ART_STOCK_MAX)
LLENA_CALREL Nulo_Valor0(artloc_llave!ART_CALIDAD)
LLENA_RELACION Nulo_Valor0(artloc_llave!ART_CODART2)
txtcodigo2.Text = Nulo_Valor0(artloc_llave!ART_CODART2)
art_situacion.Value = Val((artloc_llave!art_situacion))
frmARTI.txtcolor.Text = Trim(Nulo_Valors(artloc_llave!art_cuenta_contab))
frmARTI.txtmarca.Text = Trim(Nulo_Valors(artloc_llave!art_cuenta_contab_c))
If Nulo_Valors(artloc_llave!ART_EX_IGV) = "A" Then
   exigv.Value = 1
Else
 exigv.Value = 0
End If

checambio.Value = 0
If Nulo_Valors(artloc_llave!art_flag_cambio) = "A" Then
   checambio.Value = 1
End If

cheservi(0).Value = False
cheservi(1).Value = False
cheservi(2).Value = False
If Nulo_Valors(artloc_llave!art_flag_stock) = "M" Then
 cheservi(0).Value = True
ElseIf Nulo_Valors(artloc_llave!art_flag_stock) = "P" Then
 cheservi(1).Value = True
ElseIf Nulo_Valors(artloc_llave!art_flag_stock) = "S" Then
 cheservi(2).Value = True
End If
If Trim(gridrel.TextMatrix(1, 0)) = "" Then
 cmbcal.Visible = True
 lblcal.Visible = True
 cmdAdd.Visible = True
 cmdquitar.Visible = False
Else
 cmbcal.Visible = False
 lblcal.Visible = False
 cmdAdd.Visible = False
 cmdquitar.Visible = True
End If

If LK_EMP = "HER" Then
  txtpor1.Text = Nulo_Valor0(artloc_llave!ART_POR1)
  txtpor2.Text = Nulo_Valor0(artloc_llave!ART_POR2)
  txtpor3.Text = Nulo_Valor0(artloc_llave!ART_POR3)
  txtpor4.Text = Nulo_Valor0(artloc_llave!ART_POR4)
  txtpor5.Text = Nulo_Valor0(artloc_llave!ART_POR5)
  txtpor6.Text = Nulo_Valor0(artloc_llave!ART_POR6)
  picfoto.Visible = True
   
 
End If
If Trim(frmARTI.DS.Text) = "S" Then
 cmddolares.Caption = "Lista de Precios en S/. (Nuevos Soles)"
 cmddolares.Tag = "S"
Else
 cmddolares.Caption = "Lista de Precios en US$. (Dolares Americanos)"
 cmddolares.Tag = "D"
End If
If (updatePrecios = False) Then
 llena_pre Trim(frmARTI.DS.Text)
End If
If LK_EMP = "CAM" Then
  PROD_PROC
End If
frmARTI.SSTab1.tab = 0
VAR_ACTIVAR = 0
End Sub
Public Sub LIMPIA_ARTI()
Dim i As Integer
txtMin.Text = ""
txtMax.Text = ""
frmARTI.txt_alterno.Text = ""
frmARTI.Txt_key.Text = ""
frmARTI.txtnombre.Text = ""
DS.ListIndex = -1
lblUnidad.Caption = ""
decimales.ListIndex = -1
CmbCalidad.ListIndex = -1
art_familia.ListIndex = -1
art_subfam.ListIndex = -1
art_grupo.ListIndex = -1
art_codpro.ListIndex = -1
art_linea.ListIndex = -1
art_numero.ListIndex = -1
art_marca.ListIndex = -1
txtcospro.Text = ""
tcospro.Text = ""
DS.ListIndex = -1
art_plancha.ListIndex = -1
decimales.Text = ""

If (updatePrecios = False) Then
 grid_unid.Clear
 grid_unid.Cols = 1
 grid_unid.Rows = 1
End If

frmARTI.SSTab1.tab = 0
cheservi(0).Value = False
cheservi(1).Value = False
cheservi(2).Value = False
checambio.Value = 0
txtpeso.Text = ""
txtlitro.Text = ""
txtcolor.Text = ""
txtmarca.Text = ""
gridrel.Clear
  txtpor1.Text = ""
  txtpor2.Text = ""
  txtpor3.Text = ""
  txtpor4.Text = ""
  txtpor5.Text = ""
  txtpor6.Text = ""
  txtcodigo2.Text = ""
  exigv.Value = 0
  
VAR_ACTIVAR = 0
art_situacion.Value = 0
lblcospro.Caption = ""

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.tab = 1 And txtMin.Enabled = True Then
 ' If LK_EMP = "HER" Then
    Azul txtpor1, txtpor1
 ' Else
 '   Azul txtMin, txtMin
 ' End If
ElseIf SSTab1.tab = 0 And art_familia.Enabled = True Then
  art_familia.SetFocus
End If
End Sub


Private Sub SSTab1_GotFocus()
If ListView1.Visible Then
  ListView1.Visible = False
  Txt_key.Text = ""
End If
If SSTab1.tab = 1 And txtpor1.Enabled = True Then
  If LK_EMP = "HER" Then
    Azul txtpor1, txtpor1
  End If
ElseIf SSTab1.tab = 0 And art_familia.Enabled = True Then
  art_familia.SetFocus
End If

End Sub

Private Sub txtcospro_GotFocus()
    Azul txtcospro, txtcospro
    frmARTI.F14.Visible = False
End Sub

Private Sub txtcospro_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL txtcospro, KeyAscii
End Sub

Private Sub txtlitro_Change()
  If grid_unid.Rows > 1 Then
    grid_unid.TextMatrix(grid_unid.Row, 28) = Val(txtlitro.Text)
  End If

End Sub

Private Sub txtlitro_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL txtlitro, KeyAscii
If KeyAscii = 13 Then
  grid_unid.SetFocus
End If

End Sub

Private Sub txtMax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Left(cmdAgregar.Caption, 2) = "&G" Then
        cmdAgregar.SetFocus
    Else
        cmdModificar.SetFocus
    End If
End If
End Sub

Private Sub txtMin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Azul txtMax, txtMax
End If
End Sub

Private Sub txtnombre_GotFocus()
'    Azul txtnombre, txtnombre
End Sub

Private Sub txtnombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
  KeyAscii = 0
  Exit Sub
End If
If KeyAscii = 13 Then
    If loc_tipo = "V" Then
    If frmARTI.SSTab1.tab = 1 Then
       If LK_EMP = "HER" Then
        Azul txtpor1, txtpor1
        Exit Sub
       End If
       art_situacion.SetFocus
       Exit Sub
    End If
       If art_familia.Enabled Then art_familia.SetFocus
        DoEvents
        SendKeys "%{UP}"
        Exit Sub
    End If
End If
  
End Sub

Public Sub GRABAR_ARTI()
Dim wSTOCK  As Currency
Dim ws_igv As Currency
Dim i As Integer
Dim WS_IMPORTE As Currency
Dim WS_FLAG_UNIDAD As Integer
Dim WORIGINAL As Currency
Dim walterno As String
Dim wnombre As String
Dim WCODART2 As Currency
Dim ws_codcia As String * 2
Dim xcuenta As Integer
WS_FLAG_UNIDAD = 0
WS_IMPORTE = 0
ws_igv = 0
WORIGINAL = Val(frmARTI.Txt_key.Text)
If LK_FLAG_ORIGINAL = "A" Then
 walterno = frmARTI.Txt_key.Text
Else
 walterno = Trim(frmARTI.txt_alterno.Text)
End If
wnombre = frmARTI.txtnombre.Text
wfoto = frmARTI.picfoto.Picture
WCALIDAD = Val(Right(CmbCalidad.Text, 3))
WCODART2 = Val(txtcodigo2.Text)
PUB_CODCIA = pu_codcia 'LK_CODCIA
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If

If VAR_NEWCAL = 1 Then
  WORIGINAL = LOC_ORIGINAL
  walterno = LOC_ALTERNO
  wnombre = LOC_NOMBRE
  WCALIDAD = LOC_CALIDAD
  WCODART2 = 0
  GoTo IR_GRABA1
End If
If Left(cmdModificar.Caption, 2) = "&G" Then
   artloc_llave.Edit
   arm_llave.Edit
   arm_llave!ARM_FECHA_ULT = frmARTI.txtfechault.Text
   arm_llave.Update
Else
IR_GRABA1:
    artloc_llave.AddNew
    artloc_llave!ART_KEY = WORIGINAL
    If LK_FLAG_ORIGINAL = "A" Then
     artloc_llave!art_alterno = Trim(Str(WORIGINAL))
    Else
     artloc_llave!art_alterno = walterno
    End If
    artloc_llave!ART_POR_IGV = 0
    artloc_llave!ART_ORDEN = 1
End If
If LK_FLAG_ORIGINAL = "A" Then
 artloc_llave!art_alterno = WORIGINAL ' frmARTI.Txt_key.Text
Else
  artloc_llave!art_alterno = walterno ' Trim(txt_alterno.Text)
End If
artloc_llave!art_tipo = loc_tipo
artloc_llave!art_familia = Val(Right(art_familia.Text, 5))
artloc_llave!art_subfam = Val(Right(art_subfam.Text, 5))
artloc_llave!ART_CALIDAD = WCALIDAD
artloc_llave!art_subgru = Val(Right(art_grupo.Text, 5))
artloc_llave!art_linea = Val(Right(art_linea.Text, 5))
artloc_llave!art_plancha = Val(Right(art_plancha.Text, 6))
artloc_llave!art_numero = Val(Right(art_numero.Text, 3))
artloc_llave!art_marca = Val(Right(art_marca.Text, 3))
artloc_llave!art_codclie = Val(Right(art_codpro.Text, 6))
artloc_llave!art_nombre = wnombre
artloc_llave!ART_CODCIA = pu_codcia 'PUB_CODCIA
artloc_llave!ART_DECIMALES = Val(frmARTI.decimales.Text)
artloc_llave!ART_MONEDA = Trim(frmARTI.DS.Text)
artloc_llave!art_situacion = frmARTI.art_situacion.Value
artloc_llave!ART_STOCK_MIN = Nulo_Valor0(frmARTI.txtMin.Text)
artloc_llave!ART_STOCK_MAX = Nulo_Valor0(frmARTI.txtMax.Text)
artloc_llave!ART_POR_IGV = Nulo_Valor0(frmARTI.txtcospro.Text)
artloc_llave!ART_CODART2 = WCODART2
artloc_llave!ART_EX_IGV = ""
artloc_llave!ART_CP = ""
artloc_llave!art_cuenta_contab = Nulo_Valor0(frmARTI.txtcolor.Text)
artloc_llave!art_cuenta_contab_c = Nulo_Valor0(frmARTI.txtmarca.Text)

'artloc_llave!ART_COSPRO = Nulo_Valor0(tcospro.Text)
If exigv.Value = 1 Then
  artloc_llave!ART_EX_IGV = "A"
End If

artloc_llave!art_flag_cambio = " "
If checambio.Value = 1 Then
   artloc_llave!art_flag_cambio = "A"
End If


' ICA
If Val(artloc_llave!ART_POR_IGV) <> 0 And exigv.Value = 0 Then
   artloc_llave!ART_EX_IGV = "E"
End If
artloc_llave!art_flag_stock = ""
If cheservi(0).Value = True Then
  artloc_llave!art_flag_stock = "M"
ElseIf cheservi(1).Value = True Then
  artloc_llave!art_flag_stock = "P"
ElseIf cheservi(2).Value = True Then
  artloc_llave!art_flag_stock = "S"
End If
'If LK_EMP = "HER" Then
 artloc_llave!ART_POR1 = Val(txtpor1.Text)
 artloc_llave!ART_POR2 = Val(txtpor2.Text)
 artloc_llave!ART_POR3 = Val(txtpor3.Text)
 artloc_llave!ART_POR4 = Val(txtpor4.Text)
 artloc_llave!ART_POR5 = Val(txtpor5.Text)
 artloc_llave!ART_POR6 = Val(txtpor6.Text)
'End If
artloc_llave.Update
pu_codcia = pu_codcia 'PUB_CODCIA
PUB_CODART = WORIGINAL
SQ_OPER = 1
LEER_ARM_LLAVE
'If Not arm_llave.EOF Then
' arm_llave.Edit
' arm_llave!arm_cospro = Nulo_Valor0(tcospro.Text)
' arm_llave.Update
'End If
 
ws_codcia = pu_codcia 'LK_CODCIA
If Left(cmdModificar.Caption, 2) = "&G" Then
 GoSub IR_POR_CIA
 Exit Sub
End If
If LK_EMP_PTO = "A" Then
  If Trim(GEN!gen_ART_CIAS) <> "" Then
     xcuenta = 1
     For fila = 1 To 30
         ws_codcia = Mid(Trim(GEN!gen_ART_CIAS), xcuenta, 2)
         If Trim(ws_codcia) = "" Then Exit For
         GoSub IR_POR_CIA
         xcuenta = xcuenta + 2
     Next fila
  Else
    GoSub IR_POR_CIA
  End If
Else
  GoSub IR_POR_CIA
End If

Exit Sub


IR_POR_CIA: ' Actualiza Cias o Cia Actual

pu_codcia = ws_codcia
PUB_CODART = WORIGINAL
SQ_OPER = 2
LEER_PRE_LLAVE
If VAR_NEWCAL = 1 Then
 GoTo IR_GRABA2
End If
If Left(cmdModificar.Caption, 2) = "&G" Then
    fila = 0
    Flag_Inicial = "A"
    Do Until pre_mayor.EOF
       fila = fila + 1
      If fila >= grid_unid.Rows Then
        pre_mayor.Delete
        GoTo OTRO
      End If
      If pre_mayor!pre_secuencia <> Val(grid_unid.TextMatrix(fila, 15)) Then
        pre_mayor.Delete
        fila = fila - 1
        GoTo OTRO
      End If
       pre_mayor.Edit
       pre_mayor!pre_secuencia = Val(grid_unid.TextMatrix(fila, 15))
       pre_mayor!pre_unidad = Left(grid_unid.TextMatrix(fila, 0), 15)
       pre_mayor!PRE_EQUIV = Val(grid_unid.TextMatrix(fila, 1))
       pre_mayor!PRE_COSTO = Val((grid_unid.TextMatrix(fila, 3)) / ((100 + LK_IGV) / 100))  'agregado gts
       
       pre_mayor!pre_pre11 = Val(grid_unid.TextMatrix(fila, 16))
       pre_mayor!PRE_PRE22 = Val(grid_unid.TextMatrix(fila, 17))
       pre_mayor!PRE_PRE33 = Val(grid_unid.TextMatrix(fila, 18))
       pre_mayor!PRE_PRE44 = Val(grid_unid.TextMatrix(fila, 19))
       pre_mayor!PRE_PRE55 = Val(grid_unid.TextMatrix(fila, 20))
       ''***********************************PRECIOS6
       pre_mayor!PRE_PRE66 = Val(grid_unid.TextMatrix(fila, 31))
       ''***********************************
       pre_mayor!PRE_PRE1 = Val(grid_unid.TextMatrix(fila, 21))
       pre_mayor!PRE_PRE2 = Val(grid_unid.TextMatrix(fila, 22))
       pre_mayor!PRE_PRE3 = Val(grid_unid.TextMatrix(fila, 23))
       pre_mayor!PRE_PRE4 = Val(grid_unid.TextMatrix(fila, 24))
       pre_mayor!PRE_PRE5 = Val(grid_unid.TextMatrix(fila, 25))
       ''***********************************PRECIOS6
       pre_mayor!PRE_PRE6 = Val(grid_unid.TextMatrix(fila, 32))
       ''***********************************
       pre_mayor!PRE_FLAG_UNIDAD = grid_unid.TextMatrix(fila, 14)
       pre_mayor!pre_PESO = Val(grid_unid.TextMatrix(fila, 26))
       pre_mayor!PRE_LITRO = Val(grid_unid.TextMatrix(fila, 28))
       
       pre_mayor.Update
OTRO:
       pre_mayor.MoveNext
    Loop
    If fila > grid_unid.Rows - 1 Then
    
    ElseIf fila <> grid_unid.Rows - 1 Then
     VARFILA = fila + 1
      GoTo AGREGA
    End If

Else
IR_GRABA2:
    VARFILA = 1
AGREGA:
    fila = 0
    Flag_Inicial = "A"
    For fila = VARFILA To grid_unid.Rows - 1
       pre_mayor.AddNew
       pre_mayor!PRE_CODCIA = ws_codcia
       pre_mayor!PRE_codart = WORIGINAL
       pre_mayor!pre_secuencia = fila - 1
       pre_mayor!pre_unidad = Left(grid_unid.TextMatrix(fila, 0), 15)
       pre_mayor!PRE_EQUIV = Val(grid_unid.TextMatrix(fila, 1))
       pre_mayor!PRE_COSTO = Val(grid_unid.TextMatrix(fila, 3))
       
       pre_mayor!pre_pre11 = Val(grid_unid.TextMatrix(fila, 16))
       pre_mayor!PRE_PRE22 = Val(grid_unid.TextMatrix(fila, 17))
       pre_mayor!PRE_PRE33 = Val(grid_unid.TextMatrix(fila, 18))
       pre_mayor!PRE_PRE44 = Val(grid_unid.TextMatrix(fila, 19))
       pre_mayor!PRE_PRE55 = Val(grid_unid.TextMatrix(fila, 20))
       ''**************************************PRECIO6
       pre_mayor!PRE_PRE66 = Val(grid_unid.TextMatrix(fila, 31))
       ''**************************************
       pre_mayor!PRE_PRE1 = Val(grid_unid.TextMatrix(fila, 21))
       pre_mayor!PRE_PRE2 = Val(grid_unid.TextMatrix(fila, 22))
       pre_mayor!PRE_PRE3 = Val(grid_unid.TextMatrix(fila, 23))
       pre_mayor!PRE_PRE4 = Val(grid_unid.TextMatrix(fila, 24))
       pre_mayor!PRE_PRE5 = Val(grid_unid.TextMatrix(fila, 25))
       ''**************************************PRECIO6
       pre_mayor!PRE_PRE6 = Val(grid_unid.TextMatrix(fila, 32))
       ''**************************************
       pre_mayor!pre_PESO = Val(grid_unid.TextMatrix(fila, 26))
       pre_mayor!PRE_LITRO = Val(grid_unid.TextMatrix(fila, 28))
       pre_mayor!PRE_FLAG_UNIDAD = grid_unid.TextMatrix(fila, 14)
       pre_mayor.Update
    Next fila
End If
If Left(cmdAgregar.Caption, 2) = "&G" Then
    pu_codcia = ws_codcia
    PUB_CODART = PUB_KEY
    SQ_OPER = 1
    LEER_ARM_LLAVE
    If arm_llave.EOF Then
     Screen.MousePointer = 0
     'If LK_EMP = "CAM" Then
     '   wSTOCK = InputBox("STOCK :", "", "")
     'End If
     arm_llave.AddNew
     arm_llave!ARM_CODART = PUB_KEY
     arm_llave!ARM_CODCIA = ws_codcia
     arm_llave!ARM_STOCK = 0
     arm_llave!ARM_INGRESOS = 0
     arm_llave!ARM_SALIDAS = 0
     arm_llave!arm_cospro = 0
     arm_llave!arm_stock2 = 0
     arm_llave!ARM_COSTO_ULT = 0
     arm_llave!ARM_saldo_s = 0
     arm_llave!arm_saldo_s2 = 0
     arm_llave!ARM_Saldo_n = 0
     arm_llave!ARM_SALDO_N2 = 0
     arm_llave!ARM_FECHA_ULT = #1/1/1900#
     arm_llave.Update
     MENSAJE_ARTI "Articulo Nuevo en Compañia . . ."
    Else
      MsgBox "Codigo Existe en tabla: Articulo verificar ...", 48, Pub_Titulo
    End If
End If
Return
End Sub

Public Function GENERA_CODI() As Double
Dim NUMCAD, FIJO As String
Dim DIGI As String * 2
Dim i, VINT1, VINT2, VINT3, VINT4 As Double
Dim VSTR1, VSTR2, VSTR3, VSTR4 As String
Dim VFIJO As Double
Dim VVARI As Integer
Dim STRpub_cadena As String
Dim INTpub_cadena As Double

PUB_KEY = 0
SQ_OPER = 2
pu_codcia = LK_CODCIA
If LK_EMP_PTO = "A" Then
  pu_codcia = "00"
End If
LEER_ART_LLAVE

If art_mayor.EOF Then
    NUMCAD = "1"
Else
    art_mayor.MoveLast
    NUMCAD = art_mayor!ART_KEY
End If

VINT2 = 0
NUMCAD = Trim(NUMCAD)
VINT1 = Len(NUMCAD)
If NUMCAD = "1" Or NUMCAD = "2" Or NUMCAD = "0" Then
  VINT2 = 1
  VINT1 = 2
End If
If VINT1 > 1 Then
    VSTR4 = Val(Mid(NUMCAD, 1, VINT1 - 2)) + 1
End If

For i = 1 To VINT1 - 2
   VSTR1 = Mid(VSTR4, i, 1)
   VINT2 = VINT2 + Val(VSTR1)
Next i
VINT3 = VINT2 * 7

VSTR3 = Right(CStr(VINT3), 2)
If Len(VSTR3) = 1 Then
  VSTR3 = "0" & VSTR3
End If
FIJO = VSTR4
STRpub_cadena = FIJO & VSTR3
INTpub_cadena = Val(STRpub_cadena)

GENERA_CODI = INTpub_cadena

End Function

Public Function CONSIS_ARTI() As Boolean
Dim WSEND As String
If loc_tipo = "V" Then
    If Trim(art_familia.Text) = "" Then
        CONSIS_ARTI = False
        MsgBox "Falta seleccionar algun dato...", 48, Pub_Titulo
        art_familia.SetFocus
        SendKeys "%{UP}"
        GoTo ESCAPA
    End If
    If LK_FLAG_ORIGINAL <> "A" Then
     If txt_alterno.Text = "" Then
        CONSIS_ARTI = False
        MsgBox "Favor de Ingresar el Codigo Alterno...", 48, Pub_Titulo
        txt_alterno.SetFocus
        GoTo ESCAPA
     End If
    End If
    If txtnombre.Text = "" Then
        CONSIS_ARTI = False
        MsgBox "Falta Ingresar el NOMBRE...", 48, Pub_Titulo
        Azul txtnombre, txtnombre
        GoTo ESCAPA
    End If
    If Not IsNumeric(frmARTI.txtMin.Text) And Trim(txtMin.Text) <> "" Or Val(txtMin.Text) > 999999999.99 Then
        CONSIS_ARTI = False
        MsgBox "Dato Invalido Stock Minimo ", 48, Pub_Titulo
        frmARTI.SSTab1.tab = 1
        Azul txtMin, txtMin
        GoTo ESCAPA
    End If
    If Not IsNumeric(frmARTI.txtMax.Text) And Trim(txtMax.Text) <> "" Or Val(txtMax.Text) > 999999999.99 Then
        CONSIS_ARTI = False
        MsgBox "Dato Invalido Stock Maximo ", 48, Pub_Titulo
        frmARTI.SSTab1.tab = 1
        Azul txtMax, txtMax
        GoTo ESCAPA
    End If
    'If exigv.Value = 0 Then
    '   txtcospro.Text = ""
    'End If
    If Not IsNumeric(txtpor1) And Trim(txtpor1.Text) <> "" Or Val(txtpor1.Text) > 999.99 Then
        CONSIS_ARTI = False
        MsgBox "Dato Invalido % p' 1 ", 48, Pub_Titulo
        frmARTI.SSTab1.tab = 1
        Azul txtpor1, txtpor1
        GoTo ESCAPA
    ElseIf Not IsNumeric(txtpor2) And Trim(txtpor2.Text) <> "" Or Val(txtpor2.Text) > 999.99 Then
        CONSIS_ARTI = False
        MsgBox "Dato Invalido % p' 2 ", 48, Pub_Titulo
        frmARTI.SSTab1.tab = 1
        Azul txtpor2, txtpor2
        GoTo ESCAPA
    ElseIf Not IsNumeric(txtpor3) And Trim(txtpor3.Text) <> "" Or Val(txtpor3.Text) > 999.99 Then
        CONSIS_ARTI = False
        MsgBox "Dato Invalido % p' 3 ", 48, Pub_Titulo
        frmARTI.SSTab1.tab = 1
        Azul txtpor3, txtpor3
        GoTo ESCAPA
    ElseIf Not IsNumeric(txtpor4) And Trim(txtpor4.Text) <> "" Or Val(txtpor4.Text) > 999.99 Then
        CONSIS_ARTI = False
        MsgBox "Dato Invalido % p' 4 ", 48, Pub_Titulo
        frmARTI.SSTab1.tab = 1
        Azul txtpor4, txtpor4
        GoTo ESCAPA
    ElseIf Not IsNumeric(txtpor5) And Trim(txtpor5.Text) <> "" Or Val(txtpor5.Text) > 999.99 Then
        CONSIS_ARTI = False
        MsgBox "Dato Invalido % p' 5 ", 48, Pub_Titulo
        frmARTI.SSTab1.tab = 1
        Azul txtpor5, txtpor5
        GoTo ESCAPA
    ElseIf Not IsNumeric(txtpor6) And Trim(txtpor6.Text) <> "" Or Val(txtpor6.Text) > 999.99 Then
        CONSIS_ARTI = False
        MsgBox "Dato Invalido % p' 6 ", 48, Pub_Titulo
        frmARTI.SSTab1.tab = 1
        Azul txtpor6, txtpor6
        GoTo ESCAPA
    End If
End If

CONSIS_ARTI = True

Exit Function

ESCAPA:
 'msgbox WSEND, 48, pub_titulo

End Function

Public Sub MENSAJE_ARTI(TEXTO As String)
  LblMensaje.Caption = TEXTO
  PARPADEA.Enabled = True
End Sub
Public Sub SOLO_PORCEBTAJE(Optional tecla)
'CONVIERTE TODA A MAYUSCULAS LETRAS
Dim car As String, Longt As Integer
car = Chr$(tecla)
car = UCase$(Chr$(tecla))
tecla = Asc(car)
If car < "0" Or car > "9" Then
    If tecla <> 8 And tecla <> 13 And tecla <> 32 And car <> "." Then
        tecla = 0
        
    End If
End If
End Sub

Private Sub txt_key_GotFocus()
 If ListView1.Visible Then
  ListView1.Visible = False
 End If
 Txt_key.Text = ""
 frmARTI.F14.Visible = False
End Sub
Private Sub txt_key_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As Object    ' Variable FoundItem.
'mic para buscador
If KeyCode = vbKeyF1 Then
    frmbusqueda.Visible = True
    artfamilia.SetFocus
End If

If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And Txt_key.Text = "" Then
  loc_key = 1
  Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  DoEvents
  Txt_key.SelStart = Len(Txt_key.Text)
  DoEvents
fin:

End Sub
Private Sub txt_key_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As Object

If KeyAscii = 27 Then
 Txt_key.Text = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
On Error GoTo ERROR_CODIGO
  pu_codclie = Val(Txt_key.Text)
On Error GoTo 0
If Len(Txt_key.Text) = 0 Then
   Exit Sub
End If

If pu_codclie <> 0 And IsNumeric(Txt_key.Text) = True Then
   LOC_OPER = 1
   PUB_CODCIA = LK_CODCIA
   On Error GoTo ERROR_CODIGO
    PUB_KEY = pu_codclie
    LEER_LOC
   On Error GoTo 0
   If artloc_llave.EOF Then
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     Azul Txt_key, Txt_key
     GoTo fin
   Else
     If pu_codclie = 1 Then
       MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
       Azul Txt_key, Txt_key
       GoTo fin
     End If
     LLENA_ARTI 1
     BLOQUEA_TEXT frmARTI.Txt_key
     frmARTI.cmdModificar.SetFocus
     BLOQUEA_TEXT txtnombre
     cmdCancelar.Enabled = True
   End If
Else
   If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   valor = UCase(ListView1.ListItems.Item(loc_key).Text)
   If Trim(UCase(Txt_key.Text)) = Left(valor, Len(Trim(Txt_key.Text))) Then
   Else
      Exit Sub
   End If
   LLENA_ARTI 0
   BLOQUEA_TEXT frmARTI.Txt_key
   frmARTI.cmdModificar.SetFocus
   BLOQUEA_TEXT txtnombre
   cmdCancelar.Enabled = True
End If
dale:
ListView1.Visible = False
fin:
Exit Sub
ERROR_CODIGO:
MsgBox "Codigo NO Valido .... ", 48, Pub_Titulo
Azul Txt_key, Txt_key
End Sub

Private Sub txt_key_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
Dim ws_codcia As String * 2

If Len(Txt_key.Text) = 0 Or IsNumeric(Txt_key.Text) = True Then
   ListView1.Visible = False
   Exit Sub
End If
If (ListView1.Visible = False And KeyCode <> 13 Or Len(Txt_key.Text) = 1) Or (Left(Txt_key.Text, 1) = "%" And Trim(Len(Txt_key.Text)) > 1) Then
    If Txt_key.Text = "" Then Txt_key.Text = " "
    var = Asc(Txt_key.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    ElseIf var = 58 Then
       var = "A"
    Else
       var = Chr(var)
    End If
    ws_codcia = LK_CODCIA
    If LK_EMP_PTO = "A" Then
      ws_codcia = "00"
    End If
        
    numarchi = 0
    If Left(Txt_key.Text, 1) <> "%" Then
    ''  archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK ,PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS WHERE (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND ART_KEY  <> 1 and ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & Txt_key.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
        archi = "SELECT ARTI.ART_KEY,ARTI.ART_CODCIA, ARTI.ART_NOMBRE, ARTI.ART_ALTERNO, ARTICULO.ARM_STOCK, PRECIOS.PRE_EQUIV, TABLAS.TAB_NOMLARGO AS DIVI, TABLAS_1.TAB_NOMLARGO AS LINEA, PRECIOS.PRE_PRE1, PRECIOS.PRE_PRE2,ARTI.ART_CUENTA_CONTAB,ARTI.ART_CUENTA_CONTAB_C "
        archi = archi & "FROM ARTI INNER JOIN ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA INNER JOIN PRECIOS ON ARTI.ART_KEY = PRECIOS.PRE_CODART AND ARTI.ART_CODCIA = PRECIOS.PRE_CODCIA INNER JOIN TABLAS ON ARTI.ART_CODCIA = TABLAS.TAB_CODCIA AND ARTI.ART_FAMILIA = TABLAS.TAB_NUMTAB INNER JOIN TABLAS TABLAS_1 ON ARTI.ART_CODCIA = TABLAS_1.TAB_CODCIA AND ARTI.ART_SUBFAM = TABLAS_1.TAB_NUMTAB "
        archi = archi & "WHERE (TABLAS.TAB_TIPREG = 122) AND (TABLAS_1.TAB_TIPREG = 123) AND (PRECIOS.PRE_FLAG_UNIDAD = 'A') AND ARTI.ART_CODCIA = '" & LK_CODCIA & "' AND ARTI.ART_CALIDAD = 1 AND ARTI.ART_FLAG_STOCK = 'M' AND ARTI.ART_NOMBRE BETWEEN '" & Trim(Txt_key.Text) & "%' AND  '" & var & "' ORDER BY ARTI.ART_NOMBRE"
    Else
        If KeyCode = 13 Then
        archi = "SELECT ARTI.ART_KEY,ARTI.ART_CODCIA, ARTI.ART_NOMBRE, ARTI.ART_ALTERNO, ARTICULO.ARM_STOCK, PRECIOS.PRE_EQUIV, TABLAS.TAB_NOMLARGO AS DIVI, TABLAS_1.TAB_NOMLARGO AS LINEA, PRECIOS.PRE_PRE1, PRECIOS.PRE_PRE2,ARTI.ART_CUENTA_CONTAB,ARTI.ART_CUENTA_CONTAB_C "
        archi = archi & "FROM ARTI INNER JOIN ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA INNER JOIN PRECIOS ON ARTI.ART_KEY = PRECIOS.PRE_CODART AND ARTI.ART_CODCIA = PRECIOS.PRE_CODCIA INNER JOIN TABLAS ON ARTI.ART_CODCIA = TABLAS.TAB_CODCIA AND ARTI.ART_FAMILIA = TABLAS.TAB_NUMTAB INNER JOIN TABLAS TABLAS_1 ON ARTI.ART_CODCIA = TABLAS_1.TAB_CODCIA AND ARTI.ART_SUBFAM = TABLAS_1.TAB_NUMTAB "
        archi = archi & "WHERE (TABLAS.TAB_TIPREG = 122) AND (TABLAS_1.TAB_TIPREG = 123) AND (PRECIOS.PRE_FLAG_UNIDAD = 'A') AND ARTI.ART_CODCIA = '" & LK_CODCIA & "' AND ARTI.ART_CALIDAD = 1 AND ARTI.ART_FLAG_STOCK = 'M' AND ARTI.ART_NOMBRE like '" & Trim(Txt_key.Text) & "%' ORDER BY ARTI.ART_NOMBRE"
        Else
            Exit Sub
        End If
    End If
    PROC_LISVIEW ListView1, 3000
    loc_key = 0
    If ListView1.Visible Then
    loc_key = 1
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As Object     ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(Txt_key.Text), lvwText, , lvwPartial)
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

Public Function CONSIS_UNIDAD() As Boolean
Dim i As Integer
Dim QUIEN As String
QUIEN = 0
For fila = 1 To grid_unid.Rows - 1
   If Trim(grid_unid.TextMatrix(fila, 0)) <> "" And Val(grid_unid.TextMatrix(fila, 1)) <> 0 Then
   Else
    MsgBox "La ultima unidad ingresada no procede ...", 48, Pub_Titulo
'    grid_unid.Rows = 2
    CONSIS_UNIDAD = False
    Exit Function
   End If
   If Val(grid_unid.TextMatrix(fila, 4)) > 999.99 Or Val(grid_unid.TextMatrix(fila, 6)) > 999.99 Or Val(grid_unid.TextMatrix(fila, 8)) > 999.99 Or Val(grid_unid.TextMatrix(fila, 10)) > 999.99 Or Val(grid_unid.TextMatrix(fila, 12)) > 999.99 Then
'     MsgBox " El Procentaje debe ser menor o igual a  999.99 .", 48, Pub_Titulo
'     CONSIS_UNIDAD = False
'     Exit Function
   End If
   If Val(grid_unid.TextMatrix(fila, 4)) < -999.99 Or Val(grid_unid.TextMatrix(fila, 6)) < -999.99 Or Val(grid_unid.TextMatrix(fila, 8)) < -999.99 Or Val(grid_unid.TextMatrix(fila, 10)) < -999.99 Or Val(grid_unid.TextMatrix(fila, 12)) < -999.99 Then
'     MsgBox " El Procentaje debe ser menor o igual a  -999.99 .", 48, Pub_Titulo
'     CONSIS_UNIDAD = False
'     Exit Function
   End If
Next fila
CONSIS_UNIDAD = True
End Function


' sigue todo para el grid
Public Sub CABEZA_UNID()
Flag_Inicial = "A"
grid_unid.Cols = 33
grid_unid.Rows = 2
grid_unid.FixedCols = 0
grid_unid.FixedRows = 1
grid_unid.ColWidth(0) = 1300 ' unidad
grid_unid.ColWidth(1) = 800  ' equivalencia
grid_unid.ColWidth(2) = 1  'c. Repos.
grid_unid.ColWidth(3) = 800  ' cos. base
grid_unid.ColWidth(4) = 650   ' % 1
grid_unid.ColWidth(5) = 750   ' p 1
grid_unid.ColWidth(6) = 650   ' % 2
grid_unid.ColWidth(7) = 750   ' p 2
grid_unid.ColWidth(8) = 650   ' % 3
grid_unid.ColWidth(9) = 750   ' p 3
grid_unid.ColWidth(10) = 650   ' % 4
grid_unid.ColWidth(11) = 750   ' p 4
grid_unid.ColWidth(12) = 650   ' % 5
grid_unid.ColWidth(13) = 750   ' p 5
grid_unid.ColWidth(14) = 1   ' FLAG
grid_unid.ColWidth(15) = 1   ' SECUENCIA
grid_unid.ColWidth(16) = 1   ' Guarda p.digitado 11(S)
grid_unid.ColWidth(17) = 1   ' Guarda p.digitado 22(S)
grid_unid.ColWidth(18) = 1   ' Guarda p.digitado 33(S)
grid_unid.ColWidth(19) = 1   ' Guarda p.digitado 44(S)
grid_unid.ColWidth(20) = 1   ' Guarda p.digitado 55(S)
grid_unid.ColWidth(21) = 1   ' Guarda p.digitado 55(D)
grid_unid.ColWidth(22) = 1   ' Guarda p.digitado 55(D)
grid_unid.ColWidth(23) = 1   ' Guarda p.digitado 55(D)
grid_unid.ColWidth(24) = 1   ' Guarda p.digitado 55(D)
grid_unid.ColWidth(25) = 1   ' Guarda p.digitado 55(D)
grid_unid.ColWidth(26) = 1   ' Guarda el peso x unidad
grid_unid.ColWidth(27) = 1   ' Guarda el costo Rep. soles
grid_unid.ColWidth(28) = 1   ' Guarda el costo Rep. soles
''*********************************PRECIO6
grid_unid.ColWidth(29) = 650   ' % 5
grid_unid.ColWidth(30) = 750   ' p 5
grid_unid.ColWidth(31) = 1   ' Guarda p.digitado 66(S)
grid_unid.ColWidth(32) = 1   ' Guarda p.digitado 66(D)
''*********************************
grid_unid.TextMatrix(0, 0) = "Unidades"
grid_unid.TextMatrix(0, 1) = "Equiv."
grid_unid.TextMatrix(0, 2) = "C.Rep."
grid_unid.TextMatrix(0, 3) = "REP.C/IGV"
grid_unid.TextMatrix(0, 4) = "( % )  "
grid_unid.TextMatrix(0, 5) = "  Valor"
grid_unid.TextMatrix(0, 6) = "( % )  "
grid_unid.TextMatrix(0, 7) = "  Valor"
grid_unid.TextMatrix(0, 8) = "( % )  "
grid_unid.TextMatrix(0, 9) = "  Valor"
grid_unid.TextMatrix(0, 10) = "( % )  "
grid_unid.TextMatrix(0, 11) = "  Valor"
grid_unid.TextMatrix(0, 12) = "( % )  "
grid_unid.TextMatrix(0, 13) = "  Valor"
''*********************************PRECIO6
grid_unid.TextMatrix(0, 29) = "( % )  "
grid_unid.TextMatrix(0, 30) = "  Valor"
''*********************************

Flag_Inicial = ""
End Sub

Public Sub ElGrid_Click(wsGrid As MSFlexGrid, wsTexto As TextBox)
If wsGrid.CellWidth < 0 Then
 Exit Sub
End If
wsTexto.Left = wsGrid.Left + wsGrid.CellLeft
wsTexto.Width = wsGrid.CellWidth
wsTexto.Top = wsGrid.Top + wsGrid.CellTop
wsTexto.Tag = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
wsTexto.Text = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
wsTexto.Visible = False

End Sub
Public Sub ElGrid_EnterCell(wsGrid As MSFlexGrid, wsTexto As TextBox, Optional Bloq1, Optional Bloq2, Optional Bloq3, Optional Bloq4, Optional Bloq5)
If wsGrid.CellWidth < 0 Then
 Exit Sub
End If
If wsGrid.COL = 0 Then
 wsTexto.MaxLength = 20
Else
 wsTexto.MaxLength = 9
End If

'wsGrid.CellFontBold = True

wsGrid.CellBackColor = QBColor(7)
'wsGrid.CellForeColor = QBColor(15)

wsTexto.Left = wsGrid.Left + wsGrid.CellLeft
wsTexto.Width = wsGrid.CellWidth
wsTexto.Top = wsGrid.Top + wsGrid.CellTop
wsTexto.Tag = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
wsTexto.Text = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
Flag_Bloq = ""
If Not IsMissing(Bloq1) Then
  If wsGrid.COL = Bloq1 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq2) Then
  If wsGrid.COL = Bloq2 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq3) Then
  If wsGrid.COL = Bloq3 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq4) Then
  If wsGrid.COL = Bloq4 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq5) Then
  If wsGrid.COL = Bloq5 Then
    Flag_Bloq = "A"
  End If
End If

End Sub
Public Sub ElGrid_KeyDown(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyCode)
Flag_F2 = ""
If Flag_Bloq = "A" Then
  wsKeyCode = 0
  Exit Sub
End If

'If wsKeyCode <> 113 Then
 Exit Sub
'End If
If wsTexto.Visible = False Then
  Flag_F2 = "A"
  ElGrid_DblClick wsGrid, wsTexto
End If
End Sub
Public Sub ElGrid_KeyPress(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyAscii, Optional SaltaCol)
If wsKeyAscii = 27 Then
 Exit Sub
End If
If wsKeyAscii = 9 Or wsKeyAscii = 13 Then
  If Not IsMissing(SaltaCol) Then
    If wsGrid.COL = SaltaCol And wsGrid.Row <> wsGrid.Rows - 1 Then
       wsGrid.Row = wsGrid.Row + 1
       wsGrid.COL = wsGrid.FixedCols
       Exit Sub
    End If
    
  End If
  If wsGrid.COL <> wsGrid.Cols - 1 Then
    If wsGrid.COL = 1 Then
     wsGrid.COL = wsGrid.COL + 2
    ElseIf wsGrid.COL > 13 Then
     wsGrid.COL = wsGrid.COL + 1
    End If
  End If
  Exit Sub
End If
If Flag_Bloq = "A" Then
 wsKeyAscii = 0
 Exit Sub
End If

Dim cade
'wsTexto.FontBold = True

' wsTexto.ForeColor = QBColor(1)
wsTexto.Text = ""
wsTexto.Visible = True
cade = UCase(Chr(wsKeyAscii))
'wsTexto.text = cade
If wsTexto.Enabled = True And wsTexto.Visible = True Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
End If
Flag_Change = "A"
'cade = Chr(wsKeyAscii)
SendKeys cade, True
wsTexto.SelStart = Len(wsTexto)

End Sub
Private Sub ElGrid_LeaveCell(wsGrid As MSFlexGrid, wsTexto As TextBox)
If Flag_Consis = "A" Then
 
 'wsTexto.ForeColor = QBColor(12)
 wsTexto.Visible = True
 If wsTexto.Enabled = True And wsTexto.Visible = True Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
 End If
 Exit Sub
End If
'If Left(Trim(wsGrid.text), 1) = "-" Then
 'wsGrid.CellForeColor = QBColor(12)
 'wsGrid.CellBackColor = QBColor(15)
'Else
 wsGrid.CellBackColor = QBColor(15)
'wsGrid.CellForeColor = QBColor(0)
'End If
'wsGrid.CellFontBold = False
End Sub
Private Sub ElGrid_DblClick(wsGrid As MSFlexGrid, wsTexto As TextBox)
If Flag_Bloq = "A" Then
  Exit Sub
End If
wsTexto.FontBold = True
'wsTexto.ForeColor = QBColor(12)
wsTexto.Visible = True
If wsTexto.Enabled = True And wsTexto.Visible = True Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
End If
End Sub
Private Sub ElGrid_GotFocus(wsGrid As MSFlexGrid, wsTexto As TextBox)
ElGrid_Click wsGrid, wsTexto
End Sub
Private Sub TEXTO_LosFocus(wsGrid As MSFlexGrid, wsTexto As TextBox)
ElGrid_Click wsGrid, wsTexto
End Sub
Public Sub TEXTO_KeyDown(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyCode As Integer, Optional SaltaCol)
If wsKeyCode = 40 Or wsKeyCode = 37 Or wsKeyCode = 39 Or wsKeyCode = 38 Or wsKeyCode = 13 Then
 If Flag_F2 = "A" Then
   Exit Sub
 End If
 If Flag_Consis = "A" Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
   Beep
   Exit Sub
 End If
 Flag_Change = ""
 If wsGrid.COL <> 0 Then
  If Trim(wsTexto.Text) = "." Or Trim(wsTexto.Text) = "" Then
   wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL) = "0.00"
  Else
   If wsGrid.COL = 3 Then
    wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL) = Format(wsTexto.Text, "0.0000")
   Else
    wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL) = Format(wsTexto.Text, "0.0000")
   End If
   Select Case wsGrid.COL
     Case 5
       If frmARTI.cmddolares.Tag = "D" Then
         wsGrid.TextMatrix(wsGrid.Row, 16) = Format((wsTexto.Text) * Val(grid_unid.TextMatrix(grid_unid.Row, 1)), "0.0000")
         Else
         wsGrid.TextMatrix(wsGrid.Row, 21) = Format((wsTexto.Text) * Val(grid_unid.TextMatrix(grid_unid.Row, 1)), "0.0000")
         End If
       
     Case 7
       If frmARTI.cmddolares.Tag = "D" Then
         wsGrid.TextMatrix(wsGrid.Row, 17) = Format((wsTexto.Text) * Val(grid_unid.TextMatrix(grid_unid.Row, 1)), "0.0000")
       Else
         wsGrid.TextMatrix(wsGrid.Row, 22) = Format((wsTexto.Text) * Val(grid_unid.TextMatrix(grid_unid.Row, 1)), "0.0000")
       End If
     Case 9
       If frmARTI.cmddolares.Tag = "D" Then
         wsGrid.TextMatrix(wsGrid.Row, 18) = Format(wsTexto.Text, "0.0000")
       Else
         wsGrid.TextMatrix(wsGrid.Row, 23) = Format(wsTexto.Text, "0.0000")
       End If
     Case 11
       If frmARTI.cmddolares.Tag = "D" Then
         wsGrid.TextMatrix(wsGrid.Row, 19) = Format(wsTexto.Text, "0.0000")
       Else
         wsGrid.TextMatrix(wsGrid.Row, 24) = Format(wsTexto.Text, "0.0000")
       End If
     Case 13
       If frmARTI.cmddolares.Tag = "D" Then
         wsGrid.TextMatrix(wsGrid.Row, 20) = Format(wsTexto.Text, "0.0000")
       Else
         wsGrid.TextMatrix(wsGrid.Row, 25) = Format(wsTexto.Text, "0.0000")
       End If
    ''*********************************PRECIO6
     Case 30
       If frmARTI.cmddolares.Tag = "D" Then
         wsGrid.TextMatrix(wsGrid.Row, 31) = Format(wsTexto.Text, "0.0000")
       Else
         wsGrid.TextMatrix(wsGrid.Row, 32) = Format(wsTexto.Text, "0.0000")
       End If
    ''**********************************************
   End Select
  End If
 Else
   wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL) = wsTexto.Text
 End If
 Flag_Bloq = ""
 wsTexto.Visible = False
 If wsKeyCode = 40 Then ' ABAJO
  If wsGrid.Row <> wsGrid.Rows - 1 Then
     wsGrid.Row = wsGrid.Row + 1
  End If
 End If
 If wsKeyCode = 38 Then ' arriba
  If wsGrid.Row <> wsGrid.FixedRows Then
     wsGrid.Row = wsGrid.Row - 1
  End If
 End If
 If wsKeyCode = 37 Then ' isquierda
  If wsGrid.COL <> wsGrid.FixedCols Then
     If wsGrid.COL = 3 Then
        wsGrid.COL = wsGrid.COL - 2
     Else
        wsGrid.COL = wsGrid.COL - 1
     End If
  End If
 End If
 If wsKeyCode = 39 Then ' derecha
  If Not IsMissing(SaltaCol) Then
     If wsGrid.COL = SaltaCol Then
        If wsGrid.Row <> wsGrid.Rows - 1 Then
          wsGrid.Row = wsGrid.Row + 1
          wsGrid.COL = wsGrid.FixedCols
          GoTo wsfinal
        ElseIf wsGrid.Row = wsGrid.Rows - 1 And wsGrid.COL = wsGrid.Cols - 1 Then
         If Trim(wsGrid.TextMatrix(wsGrid.Row, 0)) <> "" And Val(wsGrid.TextMatrix(wsGrid.Row, 1)) <> 0 And Val(wsGrid.TextMatrix(wsGrid.Row, 3)) <> 0 Then
          ' wsGrid.Rows = wsGrid.Rows + 1
           'wsGrid.Row = wsGrid.Row + 1
         '  wsGrid.Col = wsGrid.FixedCols
         '  GoTo wsfinal
         Else
           wsGrid.COL = wsGrid.FixedCols
          GoTo wsfinal
         End If
        End If
     ElseIf wsGrid.Row = wsGrid.Rows - 1 And wsGrid.COL = wsGrid.Cols - 1 Then
        wsGrid.COL = wsGrid.FixedCols
        GoTo wsfinal
     End If
  End If
  If wsGrid.COL <> wsGrid.Cols - 1 Then
     If wsGrid.COL = 1 Then
         wsGrid.COL = wsGrid.COL + 2
     ElseIf wsGrid.COL >= 13 Then
         wsGrid.COL = wsGrid.FixedCols
         GoTo wsfinal
     Else
         If wsGrid.COL = 5 And LK_EMP = "3AA" Then
            wsGrid.COL = 13
         Else
           wsGrid.COL = wsGrid.COL + 1
         End If
     End If
  End If
 End If
wsfinal:
 'wsTexto.FontBold = False
 'wsTexto.ForeColor = QBColor(0)
 wsTexto.Text = ""
 wsGrid.SetFocus
End If
'Exit Sub

End Sub

Public Sub TEXTO_KeyPress(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyAscii As Integer, Optional SaltaCol, Optional ConsisCol1, Optional ConsisVal1, Optional ConsisCol2, Optional ConsisVal2, Optional ConsisCol3, Optional ConsisVal3, Optional ConsisCol4, Optional ConsisVal4, Optional ConsisCol5, Optional ConsisVal5, Optional ConsisCol6, Optional ConsisVal6, Optional ConsisCol7, Optional ConsisVal7, Optional ConsisCol8, Optional ConsisVal8, Optional ConsisCol9, Optional ConsisVal9, Optional ConsisCol10, Optional ConsisVal10, Optional ConsisCol11, Optional ConsisVal11, Optional ConsisCol12, Optional ConsisVal12)
If wsKeyAscii = 13 Or wsKeyAscii = 9 Then
  Flag_F2 = ""
  TEXTO_KeyDown wsGrid, wsTexto, 39, SaltaCol
  Exit Sub
End If
If wsKeyAscii = 27 Then
  ElGrid_Click wsGrid, wsTexto
  Flag_Change = "A"
  wsGrid.SetFocus
End If
If Not IsMissing(ConsisCol1) And Not IsMissing(ConsisVal1) Then
  If wsGrid.COL = ConsisCol1 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal1, ConsisCol1
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol2) And Not IsMissing(ConsisVal2) Then
  If wsGrid.COL = ConsisCol2 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal2, ConsisCol2
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol3) And Not IsMissing(ConsisVal3) Then
  If wsGrid.COL = ConsisCol3 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal3, ConsisCol3
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol4) And Not IsMissing(ConsisVal4) Then
  If wsGrid.COL = ConsisCol4 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal4, ConsisCol4
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol5) And Not IsMissing(ConsisVal5) Then
  If wsGrid.COL = ConsisCol5 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal5, ConsisCol5
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol6) And Not IsMissing(ConsisVal6) Then
  If wsGrid.COL = ConsisCol6 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal6, ConsisCol6
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol7) And Not IsMissing(ConsisVal7) Then
  If wsGrid.COL = ConsisCol7 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal7, ConsisCol7
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol8) And Not IsMissing(ConsisVal8) Then
  If wsGrid.COL = ConsisCol8 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal8, ConsisCol8
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol9) And Not IsMissing(ConsisVal9) Then
  If wsGrid.COL = ConsisCol9 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal9, ConsisCol9
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol10) And Not IsMissing(ConsisVal10) Then
  If wsGrid.COL = ConsisCol10 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal10, ConsisCol10
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol11) And Not IsMissing(ConsisVal11) Then
  If wsGrid.COL = ConsisCol11 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal11, ConsisCol11
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol12) And Not IsMissing(ConsisVal12) Then
  If wsGrid.COL = ConsisCol12 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal12, ConsisCol12
   Exit Sub
  End If
End If

End Sub

Private Sub Consistencias(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyAscii As Integer, Optional ConsisVal, Optional ConsisCol)
  Static valor
  Dim car As String
  Flag_Consis = ""
  If ConsisVal = 2 Then ' NUMEROS CON DECIMALES
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If wsKeyAscii = 45 Then
      If wsTexto.Text <> "" Then
         Beep
         wsKeyAscii = 0
         Exit Sub
      End If
      Flag_Consis = "A"
    End If
    If wsKeyAscii = 46 Then
      If InStr(1, wsTexto.Text, ".") <> 0 Then
        Beep
        wsKeyAscii = 0
        Exit Sub
      End If
    End If
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And wsKeyAscii <> 32 And car <> "." And car <> "-" Then
          wsKeyAscii = 0
          Beep
          Exit Sub
        End If
    End If
  ElseIf ConsisVal = 1 Then ' NUMEROS ENTEROS
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And wsKeyAscii <> 32 And car <> "-" Then
          wsKeyAscii = 0
          Beep
        End If
      End If
  End If

End Sub


Public Sub CALCULAR(wsCosto As Currency)
Dim valor As Currency
Dim tC As Integer
Flag_Inicial = "A"
tC = grid_unid.COL
valor = wsCosto * (1 + (Val(grid_unid.TextMatrix(grid_unid.Row, 4)) / 100))
If valor < 0 Then
  'grid_unid.Col = 5
  'grid_unid.CellForeColor = QBColor(12)
Else
  'grid_unid.Col = 5
  'grid_unid.CellForeColor = QBColor(0)
End If
grid_unid.TextMatrix(grid_unid.Row, 5) = Format(valor, "0.00") ' PRECIO 1 REVISAR ACA GTS
valor = wsCosto * (1 + (Val(grid_unid.TextMatrix(grid_unid.Row, 22)) / 100))
If valor < 0 Then
  'grid_unid.Col = 7
  'grid_unid.CellForeColor = QBColor(12)
Else
  'grid_unid.Col = 7
  'grid_unid.CellForeColor = QBColor(0)
End If
grid_unid.TextMatrix(grid_unid.Row, 7) = Format(valor, "0.00") ' PRECIO 2
valor = wsCosto * (1 + (Val(grid_unid.TextMatrix(grid_unid.Row, 8)) / 100))
If valor < 0 Then
  'grid_unid.Col = 9
  'grid_unid.CellForeColor = QBColor(12)
Else
  'grid_unid.Col = 9
  'grid_unid.CellForeColor = QBColor(0)
End If
grid_unid.TextMatrix(grid_unid.Row, 9) = Format(valor, "0.00") ' PRECIO 3
valor = wsCosto * (1 + (Val(grid_unid.TextMatrix(grid_unid.Row, 10)) / 100))
If valor < 0 Then
  'grid_unid.Col = 11
  'grid_unid.CellForeColor = QBColor(12)
Else
  'grid_unid.Col = 11
  'grid_unid.CellForeColor = QBColor(0)
End If
grid_unid.TextMatrix(grid_unid.Row, 11) = Format(valor, "0.00") ' PRECIO 4
valor = wsCosto * (1 + (Val(grid_unid.TextMatrix(grid_unid.Row, 12)) / 100))
If valor < 0 Then
  'grid_unid.Col = 13
  'grid_unid.CellForeColor = QBColor(12)
Else
  'grid_unid.Col = 13
  'grid_unid.CellForeColor = QBColor(0)
End If
grid_unid.TextMatrix(grid_unid.Row, 13) = Format(valor, "0.00") ' PRECIO 5
''***************************************PRECIO6
valor = wsCosto * (1 + (Val(grid_unid.TextMatrix(grid_unid.Row, 29)) / 100))
If valor < 0 Then
  'grid_unid.Col = 13
  'grid_unid.CellForeColor = QBColor(12)
Else
  'grid_unid.Col = 13
  'grid_unid.CellForeColor = QBColor(0)
End If
grid_unid.TextMatrix(grid_unid.Row, 30) = Format(valor, "0.00") ' PRECIO 5

''***************************************
grid_unid.COL = tC
Flag_Inicial = ""

End Sub

Public Sub CALCULAR_OTRO(wsCosto As Currency)
Dim valor As Currency
Dim tC As Integer
Flag_Inicial = "A"
tC = grid_unid.COL
If wsCosto = 0 Then
  GoTo CERO
End If
valor = (Val(grid_unid.TextMatrix(grid_unid.Row, 5) * 100)) / Val(grid_unid.TextMatrix(grid_unid.Row, 3)) - 100
If Val(grid_unid.TextMatrix(grid_unid.Row, 5)) = 0 Then
'   grid_unid.TextMatrix(grid_unid.Row, 4) = Format(0, "0.00") ' PRECIO 1
  'grid_unid.Col = 4
  'grid_unid.CellForeColor = QBColor(12)
Else
   grid_unid.TextMatrix(grid_unid.Row, 4) = Format(valor, "0.00") ' PRECIO 1
  'grid_unid.Col = 4
  'grid_unid.CellForeColor = QBColor(0)
End If
valor = (Val(grid_unid.TextMatrix(grid_unid.Row, 7) * 100)) / Val(grid_unid.TextMatrix(grid_unid.Row, 3)) - 100
If Val(grid_unid.TextMatrix(grid_unid.Row, 7)) = 0 Then
 '  grid_unid.TextMatrix(grid_unid.Row, 6) = Format(0, "0.00") ' PRECIO 1
  'grid_unid.Col = 6
  'grid_unid.CellForeColor = QBColor(12)
Else
   grid_unid.TextMatrix(grid_unid.Row, 6) = Format(valor, "0.00") ' PRECIO 1
  'grid_unid.Col = 6
  'grid_unid.CellForeColor = QBColor(0)
End If
valor = (Val(grid_unid.TextMatrix(grid_unid.Row, 9) * 100)) / Val(grid_unid.TextMatrix(grid_unid.Row, 3)) - 100
If Val(grid_unid.TextMatrix(grid_unid.Row, 9)) = 0 Then
  ' grid_unid.TextMatrix(grid_unid.Row, 8) = Format(0, "0.00") ' PRECIO 1
  'grid_unid.Col = 8
  'grid_unid.CellForeColor = QBColor(12)
Else
  grid_unid.TextMatrix(grid_unid.Row, 8) = Format(valor, "0.00") ' PRECIO 1
  'grid_unid.Col = 8
  'grid_unid.CellForeColor = QBColor(0)
End If
valor = (Val(grid_unid.TextMatrix(grid_unid.Row, 11) * 100)) / Val(grid_unid.TextMatrix(grid_unid.Row, 3)) - 100
If Val(grid_unid.TextMatrix(grid_unid.Row, 11)) = 0 Then
  'grid_unid.Col = 10
  'grid_unid.CellForeColor = QBColor(12)
  'grid_unid.TextMatrix(grid_unid.Row, 10) = Format(0, "0.00") ' PRECIO 1
Else
 ' grid_unid.Col = 10
 ' grid_unid.CellForeColor = QBColor(0)
 grid_unid.TextMatrix(grid_unid.Row, 10) = Format(valor, "0.00") ' PRECIO 1
End If
valor = (Val(grid_unid.TextMatrix(grid_unid.Row, 13) * 100)) / Val(grid_unid.TextMatrix(grid_unid.Row, 3)) - 100
If Val(grid_unid.TextMatrix(grid_unid.Row, 13)) = 0 Then
  'grid_unid.Col = 12
  'grid_unid.CellForeColor = QBColor(12)
  'grid_unid.TextMatrix(grid_unid.Row, 12) = Format(0, "0.00") ' PRECIO 1
Else
  'grid_unid.Col = 12
 ' grid_unid.CellForeColor = QBColor(0)
  grid_unid.TextMatrix(grid_unid.Row, 12) = Format(valor, "0.00") ' PRECIO 1
End If
''*******************************************PRECIO6
valor = (Val(grid_unid.TextMatrix(grid_unid.Row, 30) * 100)) / Val(grid_unid.TextMatrix(grid_unid.Row, 3)) - 100
If Val(grid_unid.TextMatrix(grid_unid.Row, 30)) = 0 Then
  'grid_unid.Col = 12
  'grid_unid.CellForeColor = QBColor(12)
  'grid_unid.TextMatrix(grid_unid.Row, 12) = Format(0, "0.00") ' PRECIO 6
Else
  'grid_unid.Col = 12
 ' grid_unid.CellForeColor = QBColor(0)
  grid_unid.TextMatrix(grid_unid.Row, 29) = Format(valor, "0.00") ' PRECIO 6
End If
''*************************************************
grid_unid.COL = tC
Flag_Inicial = ""
Exit Sub
CERO:
   grid_unid.TextMatrix(grid_unid.Row, 5) = Format(0, "0.00") ' PRECIO 1
   grid_unid.TextMatrix(grid_unid.Row, 7) = Format(0, "0.00") ' PRECIO 2
   grid_unid.TextMatrix(grid_unid.Row, 9) = Format(0, "0.00") ' PRECIO 3
   grid_unid.TextMatrix(grid_unid.Row, 11) = Format(0, "0.00") ' PRECIO 4
   grid_unid.TextMatrix(grid_unid.Row, 13) = Format(0, "0.00") ' PRECIO 5
   ''************************************************'PRECIO6
   grid_unid.TextMatrix(grid_unid.Row, 30) = Format(0, "0.00") ' PRECIO 6
   ''************************************************'
  Flag_Inicial = ""
End Sub



Public Sub CALCULAR_POR(WSPRE As Currency, WSCOL As Integer)
Dim valor As Currency
If Val(grid_unid.TextMatrix(grid_unid.Row, 3)) <> 0 Then
  valor = (WSPRE * 100) / Val(grid_unid.TextMatrix(grid_unid.Row, 3)) - 100
Else
  valor = 0
End If
Flag_Inicial = "A"
If valor < 0 Then
  'grid_unid.Col = WSCOL - 1
  'grid_unid.CellForeColor = QBColor(12)
Else
  'grid_unid.Col = WSCOL - 1
  'grid_unid.CellForeColor = QBColor(0)
End If
grid_unid.COL = WSCOL
Flag_Inicial = ""

grid_unid.TextMatrix(grid_unid.Row, WSCOL - 1) = Format(valor, "0.0000")


End Sub

Public Sub CALCULAR_PRE(WSPOR As Currency, WSCOL As Integer)
Dim valor As Currency
valor = Val(grid_unid.TextMatrix(grid_unid.Row, 3)) * (1 + (WSPOR / 100))
Flag_Inicial = "A"
If valor < 0 Then
  'grid_unid.Col = WSCOL + 1
  'grid_unid.CellForeColor = QBColor(12)
Else
  'grid_unid.Col = WSCOL + 1
  'grid_unid.CellForeColor = QBColor(0)
End If
grid_unid.COL = WSCOL
Flag_Inicial = ""
grid_unid.TextMatrix(grid_unid.Row, WSCOL + 1) = Format(valor, "0.0000")
  Select Case WSCOL
    Case 4
       If frmARTI.cmddolares.Tag = "D" Then
         grid_unid.TextMatrix(grid_unid.Row, 16) = Format(valor, "0.0000")
       Else
         grid_unid.TextMatrix(grid_unid.Row, 21) = Format(valor, "0.0000")
       End If
    Case 6
       If frmARTI.cmddolares.Tag = "D" Then
         grid_unid.TextMatrix(grid_unid.Row, 17) = Format(valor, "0.0000")
       Else
         grid_unid.TextMatrix(grid_unid.Row, 22) = Format(valor, "0.0000")
       End If
    Case 8
       If frmARTI.cmddolares.Tag = "D" Then
         grid_unid.TextMatrix(grid_unid.Row, 18) = Format(valor, "0.0000")
       Else
         grid_unid.TextMatrix(grid_unid.Row, 23) = Format(valor, "0.0000")
       End If
    Case 10
       If frmARTI.cmddolares.Tag = "D" Then
         grid_unid.TextMatrix(grid_unid.Row, 19) = Format(valor, "0.0000")
       Else
         grid_unid.TextMatrix(grid_unid.Row, 24) = Format(valor, "0.0000")
       End If
    Case 12
       If frmARTI.cmddolares.Tag = "D" Then
         grid_unid.TextMatrix(grid_unid.Row, 20) = Format(valor, "0.0000")
       Else
         grid_unid.TextMatrix(grid_unid.Row, 25) = Format(valor, "0.0000")
       End If
''*************************************************PRECIO6
    Case 29
       If frmARTI.cmddolares.Tag = "D" Then
         grid_unid.TextMatrix(grid_unid.Row, 31) = Format(valor, "0.0000")
       Else
         grid_unid.TextMatrix(grid_unid.Row, 32) = Format(valor, "0.0000")
       End If
''*************************************************
End Select


End Sub


Private Sub grid_UNID_Click()
ElGrid_Click grid_unid, txtvar
End Sub

Private Sub grid_UNID_DblClick()
If Flag_Inicial = "A" Then
 Exit Sub
End If
If grid_unid.COL = 1 Or grid_unid.COL = 3 Then
  Exit Sub
End If
ElGrid_DblClick grid_unid, txtvar
End Sub

Private Sub grid_UNID_EnterCell()
fcomun.Refresh
txtpeso.Text = Format(grid_unid.TextMatrix(grid_unid.Row, 26), "0.00")
txtlitro.Text = Format(grid_unid.TextMatrix(grid_unid.Row, 28), "0.00")
If Flag_Inicial = "A" Then
 Exit Sub
End If
ElGrid_EnterCell grid_unid, txtvar

End Sub

Private Sub grid_UNID_GotFocus()
If Flag_Inicial = "A" Then
 Exit Sub
End If
ElGrid_GotFocus grid_unid, txtvar
'grid_UNID.Row = loc_fila
'grid_UNID.Col = loc_colum

End Sub

Private Sub grid_UNID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 32 Then
 Exit Sub
End If
If KeyCode = 46 Then
  If grid_unid.Row <> 1 Then
    If Trim(grid_unid.TextMatrix(grid_unid.Row, 0)) <> "" And Val(grid_unid.TextMatrix(grid_unid.Row, 1)) <> 0 Then
      pub_mensaje = " Eliminar la Unidad de : " & Trim(grid_unid.TextMatrix(grid_unid.Row, 0)) & " ¿Desea Continuar... ?"
      Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
      If Pub_Respuesta = vbNo Then
        grid_unid.SetFocus
        Exit Sub
      End If
      If Trim(grid_unid.TextMatrix(grid_unid.Row, 14)) = "A" Then
         grid_unid.TextMatrix(1, 14) = "A"
         lblUnidad.Caption = Trim(grid_unid.TextMatrix(1, 0))
      End If
      grid_unid.RemoveItem grid_unid.Row
      grid_unid.SetFocus
    Else
     grid_unid.RemoveItem grid_unid.Row
     grid_unid.SetFocus
     Exit Sub
    End If
   Exit Sub
  End If
End If

If KeyCode = 45 Then
  If grid_unid.Row = grid_unid.Rows - 1 Then
    If Trim(grid_unid.TextMatrix(grid_unid.Row, 0)) <> "" And Val(grid_unid.TextMatrix(grid_unid.Row, 1)) <> 0 Then
      Flag_Inicial = "A"
      grid_unid.CellBackColor = QBColor(15)
      grid_unid.Rows = grid_unid.Rows + 1
      grid_unid.Row = grid_unid.Row + 1
      grid_unid.RowHeight(grid_unid.Row) = 285
      grid_unid.TextMatrix(grid_unid.Row, 1) = "0.00"
      grid_unid.TextMatrix(grid_unid.Row, 3) = "0.00"
      grid_unid.COL = 4
      grid_unid.CellForeColor = QBColor(9)
      grid_unid.TextMatrix(grid_unid.Row, 4) = "0.00"
      grid_unid.TextMatrix(grid_unid.Row, 5) = "0.00"
      grid_unid.COL = 6
      grid_unid.CellForeColor = QBColor(9)
      grid_unid.TextMatrix(grid_unid.Row, 6) = "0.00"
      grid_unid.TextMatrix(grid_unid.Row, 7) = "0.00"
      grid_unid.COL = 8
      grid_unid.CellForeColor = QBColor(9)
      grid_unid.TextMatrix(grid_unid.Row, 8) = "0.00"
      grid_unid.TextMatrix(grid_unid.Row, 9) = "0.00"
      grid_unid.COL = 10
      grid_unid.CellForeColor = QBColor(9)
      grid_unid.TextMatrix(grid_unid.Row, 10) = "0.00"
      grid_unid.TextMatrix(grid_unid.Row, 11) = "0.00"
      grid_unid.COL = 12
      grid_unid.CellForeColor = QBColor(9)
      grid_unid.TextMatrix(grid_unid.Row, 12) = "0.00"
      grid_unid.TextMatrix(grid_unid.Row, 13) = "0.00"
''***************************************PRECIO6
      grid_unid.COL = 29
      grid_unid.CellForeColor = QBColor(9)
      grid_unid.TextMatrix(grid_unid.Row, 29) = "0.00"
      grid_unid.TextMatrix(grid_unid.Row, 30) = "0.00"
''***************************************************
      Flag_Inicial = ""
      grid_unid.COL = 0
      Exit Sub
    End If
  End If
End If
If grid_unid.COL = 1 And grid_unid.Row = 1 Then
 Exit Sub
End If

ElGrid_KeyDown grid_unid, txtvar, KeyCode
End Sub

Private Sub grid_UNID_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 And grid_unid.COL = 0 Then
 For fila = 1 To grid_unid.Rows - 1
    grid_unid.TextMatrix(fila, 14) = " "
 Next fila
 grid_unid.TextMatrix(grid_unid.Row, 14) = "A"
 lblUnidad.Caption = grid_unid.TextMatrix(grid_unid.Row, 0)
 Exit Sub
End If
If grid_unid.COL = 1 And grid_unid.Row = 1 Then
 Exit Sub
End If
'If grid_unid.Col = 3 Then Exit Sub
ElGrid_KeyPress grid_unid, txtvar, KeyAscii, 30
End Sub

Private Sub grid_UNID_LeaveCell()
If Flag_Inicial = "A" Then
 Exit Sub
End If
If Flag_Change <> "A" Then
  If grid_unid.COL = 3 Then ' costo base
    If LK_FLAG_CALCULO = "A" Then
      CALCULAR Val(grid_unid.TextMatrix(grid_unid.Row, 3))
    Else
      CALCULAR_OTRO Val(grid_unid.TextMatrix(grid_unid.Row, 3))
    End If
  End If
  If grid_unid.COL = 5 Or grid_unid.COL = 7 Or grid_unid.COL = 9 Or grid_unid.COL = 11 Or grid_unid.COL = 13 Or grid_unid.COL = 30 Then ' costo PORCENTAJE
    CALCULAR_POR Val(grid_unid.TextMatrix(grid_unid.Row, grid_unid.COL)), grid_unid.COL
  End If
  If grid_unid.COL = 4 Or grid_unid.COL = 6 Or grid_unid.COL = 8 Or grid_unid.COL = 10 Or grid_unid.COL = 12 Or grid_unid.COL = 29 Then ' costo PORCENTAJE
    CALCULAR_PRE Val(grid_unid.TextMatrix(grid_unid.Row, grid_unid.COL)), grid_unid.COL
  End If
  Flag_Change = "A"
End If
ElGrid_LeaveCell grid_unid, txtvar
End Sub

Private Sub txtpeso_Change()
  If grid_unid.Rows > 1 Then
    grid_unid.TextMatrix(grid_unid.Row, 26) = Val(txtpeso.Text)
  End If
End Sub

Private Sub txtpeso_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL txtpeso, KeyAscii
If KeyAscii = 13 Then
  grid_unid.SetFocus
End If

End Sub

Private Sub txtpor1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul txtpor2, txtpor2
End If
End Sub

Private Sub txtpor2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul txtpor3, txtpor3
End If

End Sub

Private Sub txtpor3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Azul txtpor4, txtpor4
    End If
End Sub

Private Sub txtpor4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Azul txtpor5, txtpor5
    End If
End Sub

Private Sub txtpor5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Azul txtpor6, txtpor6
    End If
End Sub
Private Sub txtpor6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Azul txtMin, txtMin
    End If
End Sub

Private Sub txtvar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 37 Or KeyCode = 39 Or KeyCode = 38 Then
 If grid_unid.COL = 4 Or grid_unid.COL = 6 Or grid_unid.COL = 8 Or grid_unid.COL = 10 Or grid_unid.COL = 12 Then
   If Val(txtvar.Text) > 999.99 Or Val(txtvar.Text) < 0 Then
     txtvar.Visible = True
     Exit Sub
    End If
  End If
  grid_unid.SetFocus
End If

If KeyCode = 40 Or KeyCode = 37 Or KeyCode = 39 Or KeyCode = 38 Or KeyCode = 13 Then
 If grid_unid.COL = 1 Then
   If Val(grid_unid.TextMatrix(grid_unid.Row, 3)) = 0 Then
    grid_unid.TextMatrix(grid_unid.Row, 3) = Format(Val(grid_unid.TextMatrix(1, 3)) * Val(txtvar.Text), "0.0000")
     If LK_FLAG_CALCULO = "A" Then
       CALCULAR Val(grid_unid.TextMatrix(grid_unid.Row, 3))
     Else
       CALCULAR_OTRO Val(grid_unid.TextMatrix(grid_unid.Row, 3))
     End If
   End If
 End If
End If

TEXTO_KeyDown grid_unid, txtvar, KeyCode, 13
If KeyCode = 13 Or KeyCode = 27 Then
 grid_unid.SetFocus
End If
End Sub

Private Sub txtvar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If grid_unid.COL = 4 Or grid_unid.COL = 6 Or grid_unid.COL = 8 Or grid_unid.COL = 10 Or grid_unid.COL = 12 Then
  If Val(txtvar.Text) > 999.99 Or Val(txtvar.Text) < 0 Then
    txtvar.Visible = True
    Exit Sub
  End If
End If
End If

TEXTO_KeyPress grid_unid, txtvar, KeyAscii, 13, 1, 2, 3, 2, 4, 2, 5, 2, 6, 2, 7, 2, 8, 2, 9, 2, 10, 2, 11, 2, 12, 2, 13, 2
End Sub

Private Sub txtvar_LostFocus()
TEXTO_LosFocus grid_unid, txtvar
End Sub

Private Sub txt_alterno_GotFocus()
If Left(cmdAgregar.Caption, 2) = "&A" Or Left(cmdModificar.Caption, 2) = "&M" Then
 Exit Sub
End If
 If ListView1.Visible Then
  ListView1.Visible = False
 End If
 Azul txt_alterno, txt_alterno
 frmARTI.F14.Visible = False
End Sub
Private Sub txt_alterno_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As Object    ' Variable FoundItem.

'mic para buscador
If KeyCode = vbKeyF1 Then
    frmbusqueda.Visible = True
    artfamilia.SetFocus
End If

If Not ListView1.Visible Or Left(cmdAgregar.Caption, 2) = "&G" Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And txt_alterno.Text = "" Then
  loc_key = 1
  Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  txt_alterno.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  DoEvents
  txt_alterno.SelStart = Len(txt_alterno.Text)
  DoEvents
fin:

End Sub
Private Sub txt_alterno_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As Object
valor = Chr(KeyAscii)
KeyAscii = Asc(UCase(valor))
valor = ""
If KeyAscii = 27 Then
  ListView1.Visible = False
 txt_alterno.Text = ""
End If
If Left(cmdAgregar.Caption, 2) = "&G" Then
 If KeyAscii = 13 Then
   txtnombre.SetFocus
   Exit Sub
 End If
 Exit Sub
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
If VAR_ACTIVAR <> 99 Then
 LOC_OPER = 1
 pu_alterno = txt_alterno.Text
 PUB_CODCIA = LK_CODCIA
 LEER_LOC
 If artloc_llave.EOF Then
   MsgBox "Codigo No Existe ...", 48, Pub_Titulo
   Azul txt_alterno, txt_alterno
   Exit Sub
 End If
 LLENA_ARTI 0
 BLOQUEA_TEXT txt_alterno
 frmARTI.cmdModificar.SetFocus
 BLOQUEA_TEXT txtnombre
 cmdCancelar.Enabled = True
 Exit Sub
End If
pu_alterno = Trim(txt_alterno.Text)
If Len(txt_alterno.Text) = 0 Then
   Exit Sub
End If
If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
  Exit Sub
End If
valor = UCase(ListView1.ListItems.Item(loc_key).Text)
If Trim(UCase(txt_alterno.Text)) = Left(valor, Len(Trim(txt_alterno.Text))) Then
Else
   Exit Sub
End If
LLENA_ARTI 0
BLOQUEA_TEXT txt_alterno
frmARTI.cmdModificar.SetFocus
BLOQUEA_TEXT txtnombre
cmdCancelar.Enabled = True
dale:
ListView1.Visible = False

fin:
End Sub

Private Sub txt_alterno_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
Dim ws_codcia As String * 2
If KeyCode = 13 Or KeyCode = 27 Then VAR_ACTIVAR = 0: Exit Sub
If Left(cmdAgregar.Caption, 2) = "&G" Then Exit Sub
If txt_alterno.Text = "*" And KeyCode = 106 Then
 VAR_ACTIVAR = 99
 Exit Sub
ElseIf txt_alterno.Text = "" Then
  ListView1.Visible = False
  VAR_ACTIVAR = 0
  Exit Sub
ElseIf txt_alterno.Text = "* " And KeyCode = 106 Then
 VAR_ACTIVAR = 99
 txt_alterno.Text = "*"
 txt_alterno.SelStart = Len(txt_alterno.Text)
 KeyCode = 0
 Exit Sub
End If

If VAR_ACTIVAR <> 99 Then
 Exit Sub
End If
If txt_alterno.Text = "*" Then
 Exit Sub
ElseIf Left(txt_alterno.Text, 1) = "*" Then
 txt_alterno.Text = Mid(txt_alterno.Text, 2, Len(txt_alterno.Text))
 txt_alterno.SelStart = Len(txt_alterno.Text)
End If
If Len(txt_alterno.Text) = 0 Or txt_alterno.Text = "" Or Left(cmdAgregar.Caption, 2) = "&G" Then
   ListView1.Visible = False
   Exit Sub
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(txt_alterno.Text) = 1 Then
    If txt_alterno.Text = "" Then txt_alterno.Text = " "
    var = Asc(txt_alterno.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    ElseIf var = 58 Then
       var = "A"
    'ElseIf var = 50 Then
    '   var = " 1"
    Else
       var = Chr(var)
    End If
    ws_codcia = LK_CODCIA
    If LK_EMP_PTO = "A" Then
    ws_codcia = "00"
    End If
    numarchi = 3
    ' archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK FROM ARTI, ARTICULO WHERE (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND  ART_KEY <> 0 AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & txt_alterno.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
    archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK ,PRE_EQUIV,    FROM ARTI, ARTICULO, PRECIOS WHERE (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND ART_KEY  <> 1 and ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & txt_alterno.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
    PROC_LISVIEW ListView1, 1000
    loc_key = 0
    If ListView1.Visible Then
     loc_key = 1
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As Object    ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(txt_alterno.Text), lvwText, , lvwPartial)
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
  End If
  Exit Sub
End If

End Sub

Public Sub PROCESO_ARTI()
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
 cade = "SELECT * FROM ARTI WHERE ART_ALTERNO = ? AND ART_CODCIA = ?  AND ART_TIPO = ? ORDER BY ART_CODCIA, ART_KEY"
Else
 cade = "SELECT * FROM ARTI WHERE ART_KEY = ? AND ART_CODCIA = ?  AND ART_TIPO = ? ORDER BY ART_CODCIA, ART_KEY"
End If
Set PSART_LOC = CN.CreateQuery("", cade)
 If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  PSART_LOC.rdoParameters(0) = " "
 Else
  PSART_LOC.rdoParameters(0) = 0
 End If
 PSART_LOC.rdoParameters(1) = " "
 PSART_LOC.rdoParameters(2) = " "
Set artloc_llave = PSART_LOC.OpenResultset(rdOpenKeyset, rdConcurValues)
End Sub

Public Sub PROCESO_CANCELAR()
    If Left(cmdAgregar.Caption, 2) = "&A" And Left(cmdModificar.Caption, 2) = "&M" Then
        frmARTI.frarelacion.Enabled = False
        LIMPIA_ARTI
        cmdCancelar.Enabled = True
        BLOQUEA_TEXT txtnombre, CmbCalidad, decimales, DS, txtcospro, art_situacion, art_linea, art_numero, art_marca, art_plancha
        BLOQUEA_TEXT art_grupo, art_familia, art_subfam, grid_unid, txtMin, txtMax, art_codpro, txtcodigo2
        BLOQUEA_TEXT cheservi(0), cheservi(1), cheservi(2), exigv, txtcospro, cmddolares, txtpeso, txtfechault, checambio, txtlitro
        BLOQUEA_TEXT txtcolor, txtmarca
        If LK_EMP = "HER" Then
           BLOQUEA_TEXT txtpor1, txtpor2, txtpor3, txtpor4, txtpor5, txtpor6
           picfoto.Visible = False
        End If
        frmARTI.SSTab1.tab = 0
        If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
          DESBLOQUEA_TEXT txt_alterno
          DESBLOQUEA_TEXT txtpath  'gts
          BLOQUEA_TEXT Txt_key
          If frmARTI.txt_alterno.Visible Then frmARTI.txt_alterno.SetFocus
        Else
          BLOQUEA_TEXT txt_alterno
          DESBLOQUEA_TEXT Txt_key
          If frmARTI.Txt_key.Visible Then frmARTI.Txt_key.SetFocus
        End If
        MANOS(0).Enabled = True
        MANOS(1).Enabled = True
        pasa = 0
        Exit Sub
    End If
    Screen.MousePointer = 11
    If Left(cmdModificar.Caption, 2) = "&G" Then
       cmdModificar.Caption = "&Modificación"
       LLENA_ARTI 1
       BLOQUEA_TEXT Txt_key
       BLOQUEA_TEXT txt_alterno
    Else
       frmARTI.frarelacion.Enabled = False
       cmdAgregar.Caption = "&Adicionar"
       LIMPIA_ARTI
       If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
          DESBLOQUEA_TEXT txt_alterno
          BLOQUEA_TEXT Txt_key
          frmARTI.txt_alterno.SetFocus
        Else
          DESBLOQUEA_TEXT Txt_key
          BLOQUEA_TEXT txt_alterno
          frmARTI.Txt_key.SetFocus
        End If
    End If
    cmdAgregar.Enabled = True
    cmdEliminar.Enabled = True
    cmdModificar.Enabled = True
    BLOQUEA_TEXT txtnombre, CmbCalidad, decimales, DS, txtcospro, art_situacion, art_linea, art_numero, art_marca, art_plancha
    BLOQUEA_TEXT art_grupo, art_familia, art_subfam, grid_unid, txtMin, txtMax, art_codpro, txtcodigo2
    BLOQUEA_TEXT cheservi(0), cheservi(1), cheservi(2), exigv, txtcospro, cmddolares, txtpeso, txtfechault, checambio, txtlitro
    BLOQUEA_TEXT txtcolor, txtmarca
    If LK_EMP = "HER" Then
        BLOQUEA_TEXT txtpor1, txtpor2, txtpor3, txtpor4, txtpor5, txtpor6
    End If
    pasa = 0
    MANOS(0).Enabled = True
    MANOS(1).Enabled = True
    MENSAJE_ARTI "Proceso Cancelado... !!!    "
    frmARTI.SSTab1.tab = 0
    Screen.MousePointer = 0
End Sub


Public Sub PROCESA_PROV()
   SQ_OPER = 2
   pu_cp = "P"
   pu_codclie = 0
   pu_codcia = LK_CODCIA
   LEER_CLI_LLAVE
   art_codpro.Clear
   Do Until cli_mayor.EOF
     art_codpro.AddItem cli_mayor!CLI_NOMBRE & String(20, " ") & Trim(CStr(cli_mayor!cli_codclie))
     cli_mayor.MoveNext
   Loop
End Sub

Public Function EXISTE_ART(WARTI As String, WCODI As String) As Boolean
Dim var
Dim tempo
tempo = Left(Trim(WARTI), Len(WARTI) - 1)
var = Asc(Right(Trim(WARTI), 1))
var = var + 1
If var = 91 Then
  var = "ZZZZZZZZ"
Else
  var = Chr(var)
End If
tempo = tempo + var
'tempo = Replace(tempo, "'", "")
archi = "SELECT * FROM ARTI WHERE  ART_KEY <> " & WCODI & " AND ART_KEY  <> 1 and ART_CODCIA = '" & LK_CODCIA & "' AND ART_NOMBRE BETWEEN '" & WARTI & "' AND  '" & tempo & "' ORDER BY ART_NOMBRE "
ListExiste.Clear
F14.Caption = "Lista de Articulos  Parecidos ... "
frmARTI.ListExiste.TextMatrix(0, 0) = "Cia"
frmARTI.ListExiste.TextMatrix(0, 1) = "Key "
frmARTI.ListExiste.TextMatrix(0, 2) = "Articulo"
If LK_FLAG_ORIGINAL <> "A" Then
 frmARTI.ListExiste.TextMatrix(0, 3) = "Codigo"
 frmARTI.ListExiste.ColAlignment(3) = 2
End If
frmARTI.ListExiste.TextMatrix(0, 4) = "SELECT"
EXISTE_ART = False
Set PSX = CN.CreateQuery("", archi)

Set X = PSX.OpenResultset(rdOpenKeyset)
X.Requery
If X.EOF Then
 frmARTI.ListExiste.Clear
 GoTo fin
End If

fila = 0
frmARTI.ListExiste.Rows = 2
Do Until X.EOF
    fila = fila + 1
    frmARTI.ListExiste.TextMatrix(fila, 0) = Nulo_Valors(X!ART_CODCIA)
    frmARTI.ListExiste.TextMatrix(fila, 1) = Nulo_Valor0(X!ART_KEY)
    frmARTI.ListExiste.TextMatrix(fila, 2) = Nulo_Valors(X!art_nombre)
    frmARTI.ListExiste.TextMatrix(fila, 4) = "1"
    If LK_FLAG_ORIGINAL <> "A" Then
     frmARTI.ListExiste.TextMatrix(fila, 3) = Trim(Nulo_Valors(X!art_alterno))
    End If
    frmARTI.ListExiste.Rows = frmARTI.ListExiste.Rows + 1
    X.MoveNext
Loop
EXISTE_ART = True
If EXISTE_ART Then
    frmARTI.ListExiste.Rows = frmARTI.ListExiste.Rows - 1
    Op(0).Value = True
    'Op(0).Enabled = False
    Op(1).Value = False
    frmARTI.F14.Visible = True
    frmARTI.ListExiste.Row = 1
    frmARTI.ListExiste.COL = 1
    frmARTI.ListExiste.SetFocus
End If
GoTo fin
Exit Function

CHECKERROR:
MsgBox Err.Description
fin:

End Function


Public Sub CABE_RELACION()
gridrel.Cols = 4
gridrel.Rows = 1
gridrel.TextMatrix(0, 0) = "Cod.Orig."
gridrel.TextMatrix(0, 1) = "Cod.Alterno"
gridrel.TextMatrix(0, 2) = "Descripción"
gridrel.TextMatrix(0, 3) = "Calidad"
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
gridrel.ColWidth(1) = 1000
Else
gridrel.ColWidth(1) = 1
End If
gridrel.ColWidth(0) = 1000
gridrel.ColWidth(2) = 3100
gridrel.ColWidth(3) = 2000

End Sub

Public Sub LLENA_RELACION(Wkey_Rela As Currency)
gridrel.Clear
CABE_RELACION
gridrel.Rows = 2
SQ_OPER = 1
pu_codcia = LK_CODCIA
If LK_EMP_PTO = "A" Then
  pu_codcia = "00"
End If
PUB_KEY = Wkey_Rela
PUB_CODCIA = LK_CODCIA
LOC_OPER = 2
LEER_LOC
If artloc_key.EOF Then
   artloc_llave.Edit
   artloc_llave!ART_CODART2 = 0
   artloc_llave.Update
   Exit Sub
End If
 If artloc_key!ART_KEY = 0 Then Exit Sub
 gridrel.RowHeight(1) = 300
 gridrel.TextMatrix(1, 0) = artloc_key!ART_KEY
 gridrel.TextMatrix(1, 1) = artloc_key!art_alterno
 gridrel.TextMatrix(1, 2) = artloc_key!art_nombre
 SQ_OPER = 1
 PUB_TIPREG = 2
 PUB_NUMTAB = artloc_key!ART_CALIDAD
 PUB_CODCIA = LK_CODCIA
 LEER_TAB_LLAVE
 If tab_llave.EOF Then
  gridrel.TextMatrix(1, 3) = ""
 Else
  gridrel.TextMatrix(1, 3) = Trim(tab_llave!tab_NOMLARGO)
 End If

End Sub

Public Sub LLENA_CALREL(wcla_actual As Integer)
Dim wa As Integer
PUB_TIPREG = 2
PUB_CODCIA = LK_CODCIA
SQ_OPER = 2
LEER_TAB_LLAVE
cmbcal.Clear
wa = 0
Do Until tab_mayor.EOF
    If tab_mayor!TAB_NUMTAB > wcla_actual Then
       cmbcal.AddItem tab_mayor!tab_NOMLARGO & String(50, " ") & tab_mayor!TAB_NUMTAB
    End If
    wa = 1
    tab_mayor.MoveNext
Loop
If cmbcal.ListCount = 0 And wa = 1 Then
  cmbcal.AddItem "<Ninguno>"
End If
If cmbcal.ListCount > 0 Then cmbcal.ListIndex = 0
End Sub

Public Sub llena_pre(wlista As String)
pu_codcia = LK_CODCIA
PUB_CODART = Val(Txt_key.Text)
SQ_OPER = 2
LEER_PRE_LLAVE
If pre_mayor.EOF Then
 MsgBox "Error de Unidades NO existe......", 48, Pub_Titulo
 Exit Sub
End If
fila = 0
Flag_Inicial = "A"
grid_unid.Clear
CABEZA_UNID
Do Until pre_mayor.EOF
   fila = fila + 1
   grid_unid.Rows = fila + 1
   grid_unid.Row = fila
   grid_unid.RowHeight(fila) = 285
   grid_unid.TextMatrix(fila, 0) = Trim(pre_mayor!pre_unidad)
   grid_unid.TextMatrix(fila, 1) = pre_mayor!PRE_EQUIV
   grid_unid.TextMatrix(fila, 2) = ""
   'grid_unid.TextMatrix(fila, 3) = Format((Nulo_Valor0(pre_mayor!PRE_COSTO) * ((100 + LK_IGV) / 100)), "0.0000")
   'AGREGADO PARA MOSTRAR EL COSTO EDITADO
   grid_unid.TextMatrix(fila, 3) = Format((Nulo_Valor0(pre_mayor!PRE_COSTO) * ((100 + LK_IGV) / 100)), "0.0000") 'agregado gts
   'grid_unid.TextMatrix(fila, 27) = Format((Nulo_Valor0(pre_mayor!PRE_COSTO) * ((100 + LK_IGV) / 100)), "0.0000")
   grid_unid.TextMatrix(fila, 27) = Nulo_Valor0(pre_mayor!PRE_COSTO)
   grid_unid.COL = 4
   grid_unid.CellForeColor = QBColor(9)
   If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then
    WSPOR = (Nulo_Valor0(pre_mayor!PRE_PRE1) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
   End If
   grid_unid.TextMatrix(fila, 4) = Format(WSPOR, "0.00")
   If wlista = "S" Then
     grid_unid.TextMatrix(fila, 5) = Nulo_Valor0(pre_mayor!PRE_PRE1)
   Else
     grid_unid.TextMatrix(fila, 5) = Nulo_Valor0(pre_mayor!pre_pre11)
   End If
   grid_unid.TextMatrix(fila, 16) = Nulo_Valor0(pre_mayor!pre_pre11)
   grid_unid.TextMatrix(fila, 21) = Nulo_Valor0(pre_mayor!PRE_PRE1)
   grid_unid.COL = 6
   grid_unid.CellForeColor = QBColor(9)
   If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then WSPOR = (Nulo_Valor0(pre_mayor!PRE_PRE2) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
   grid_unid.TextMatrix(fila, 6) = Format(WSPOR, "0.00")
   If wlista = "S" Then
    grid_unid.TextMatrix(fila, 7) = Nulo_Valor0(pre_mayor!PRE_PRE2)
   Else
    grid_unid.TextMatrix(fila, 7) = Nulo_Valor0(pre_mayor!PRE_PRE22)
   End If
   grid_unid.TextMatrix(fila, 17) = Nulo_Valor0(pre_mayor!PRE_PRE22)
   grid_unid.TextMatrix(fila, 22) = Nulo_Valor0(pre_mayor!PRE_PRE2)
   grid_unid.COL = 8
   grid_unid.CellForeColor = QBColor(9)
   If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then WSPOR = (Nulo_Valor0(pre_mayor!PRE_PRE3) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
   grid_unid.TextMatrix(fila, 8) = Format(WSPOR, "0.00")
   If wlista = "S" Then
    grid_unid.TextMatrix(fila, 9) = Nulo_Valor0(pre_mayor!PRE_PRE3)
   Else
    grid_unid.TextMatrix(fila, 9) = Nulo_Valor0(pre_mayor!PRE_PRE33)
   End If
   grid_unid.TextMatrix(fila, 18) = Nulo_Valor0(pre_mayor!PRE_PRE33)
   grid_unid.TextMatrix(fila, 23) = Nulo_Valor0(pre_mayor!PRE_PRE3)
   grid_unid.COL = 10
   grid_unid.CellForeColor = QBColor(9)
   If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then WSPOR = (Nulo_Valor0(pre_mayor!PRE_PRE4) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
   grid_unid.TextMatrix(fila, 10) = Format(WSPOR, "0.00")
   If wlista = "S" Then
     grid_unid.TextMatrix(fila, 11) = Nulo_Valor0(pre_mayor!PRE_PRE4)
   Else
     grid_unid.TextMatrix(fila, 11) = Nulo_Valor0(pre_mayor!PRE_PRE44)
   End If
   grid_unid.TextMatrix(fila, 19) = Nulo_Valor0(pre_mayor!PRE_PRE44)
   grid_unid.TextMatrix(fila, 24) = Nulo_Valor0(pre_mayor!PRE_PRE4)
   grid_unid.COL = 12
   grid_unid.CellForeColor = QBColor(9)
   If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then WSPOR = (Nulo_Valor0(pre_mayor!PRE_PRE5) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
   grid_unid.TextMatrix(fila, 12) = Format(WSPOR, "0.00")
   If wlista = "S" Then
    grid_unid.TextMatrix(fila, 13) = Nulo_Valor0(pre_mayor!PRE_PRE5)
   Else
    grid_unid.TextMatrix(fila, 13) = Nulo_Valor0(pre_mayor!PRE_PRE55)
   End If
   grid_unid.TextMatrix(fila, 20) = Nulo_Valor0(pre_mayor!PRE_PRE55)
   grid_unid.TextMatrix(fila, 25) = Nulo_Valor0(pre_mayor!PRE_PRE5)
''*******************************************************PRECIO6
   grid_unid.COL = 29
   grid_unid.CellForeColor = QBColor(9)
   If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then WSPOR = (Nulo_Valor0(pre_mayor!PRE_PRE6) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
   grid_unid.TextMatrix(fila, 29) = Format(WSPOR, "0.00")
   If wlista = "S" Then
    grid_unid.TextMatrix(fila, 30) = Nulo_Valor0(pre_mayor!PRE_PRE6)
   Else
    grid_unid.TextMatrix(fila, 30) = Nulo_Valor0(pre_mayor!PRE_PRE66)
   End If
   grid_unid.TextMatrix(fila, 31) = Nulo_Valor0(pre_mayor!PRE_PRE66)
   grid_unid.TextMatrix(fila, 32) = Nulo_Valor0(pre_mayor!PRE_PRE6)
''******************************************************
   grid_unid.TextMatrix(fila, 14) = Nulo_Valor0(pre_mayor!PRE_FLAG_UNIDAD)
   grid_unid.TextMatrix(fila, 15) = Nulo_Valor0(pre_mayor!pre_secuencia)
   grid_unid.TextMatrix(fila, 26) = Nulo_Valor0(pre_mayor!pre_PESO)
   grid_unid.TextMatrix(fila, 28) = Nulo_Valor0(pre_mayor!PRE_LITRO)
   If Nulo_Valor0(pre_mayor!PRE_FLAG_UNIDAD) = "A" Then
     lblUnidad.Caption = Trim(pre_mayor!pre_unidad)
     frmARTI.lblcospro.Caption = Format(Val(frmARTI.lblcospro.Caption) * Trim(pre_mayor!PRE_EQUIV), "###,##0.000")
   End If
   pre_mayor.MoveNext
Loop

Flag_Inicial = ""
grid_unid.Row = 1
grid_unid.COL = 0
If LK_EMP = "3AA" Then
'  cmddolares_Click
End If

End Sub

Public Sub ARTI_CERO()

PUB_KEY = 0
artloc_llave.AddNew
artloc_llave!ART_KEY = PUB_KEY
artloc_llave!art_alterno = PUB_KEY
artloc_llave!ART_POR_IGV = 0
artloc_llave!ART_ORDEN = 0
artloc_llave!art_tipo = loc_tipo
artloc_llave!art_familia = 0
artloc_llave!art_subfam = 0
artloc_llave!ART_CALIDAD = 1
artloc_llave!art_subgru = 0
artloc_llave!art_codclie = 0
artloc_llave!art_nombre = "Productos"
artloc_llave!ART_CODCIA = LK_CODCIA
artloc_llave!ART_DECIMALES = 2
artloc_llave!ART_MONEDA = "S"
artloc_llave!art_situacion = 0
artloc_llave!ART_STOCK_MIN = 0
artloc_llave!ART_STOCK_MAX = 0
artloc_llave!ART_POR_IGV = 0
artloc_llave!ART_CODART2 = 0
artloc_llave!ART_EX_IGV = 0
artloc_llave!ART_COSPRO = 0
artloc_llave!ART_EX_IGV = "A"
artloc_llave!art_flag_stock = "M"
artloc_llave!ART_POR1 = 0
artloc_llave!ART_POR2 = 0
artloc_llave!ART_POR3 = 0
artloc_llave!ART_POR4 = 0
artloc_llave!ART_POR5 = 0
artloc_llave!ART_POR6 = 0
artloc_llave.Update

pre_mayor.AddNew
pre_mayor!PRE_CODCIA = LK_CODCIA
pre_mayor!PRE_codart = PUB_KEY
pre_mayor!pre_secuencia = 1
pre_mayor!pre_unidad = "UNIDAD"
pre_mayor!PRE_EQUIV = 1

pre_mayor!pre_pre11 = 0
pre_mayor!PRE_PRE22 = 0
pre_mayor!PRE_PRE33 = 0
pre_mayor!PRE_PRE44 = 0
pre_mayor!PRE_PRE55 = 0
''******************************PRECIO6
pre_mayor!PRE_PRE66 = 0
''*************************************
pre_mayor!PRE_PRE1 = 0
pre_mayor!PRE_PRE2 = 0
pre_mayor!PRE_PRE3 = 0
pre_mayor!PRE_PRE4 = 0
pre_mayor!PRE_PRE5 = 0
''******************************PRECIO6
pre_mayor!PRE_PRE6 = 0
''*************************************
pre_mayor!pre_PESO = 0
pre_mayor!PRE_LITRO = 0
pre_mayor!PRE_FLAG_UNIDAD = "A"
pre_mayor.Update

arm_llave.AddNew
arm_llave!ARM_CODART = PUB_KEY
arm_llave!ARM_CODCIA = LK_CODCIA
arm_llave!ARM_STOCK = 0
''arm_llave!ARM_STOCKS = 0
''arm_llave!ARM_STOCKS2 = 0
''arm_llave!ARM_STOCKN = 0
''arm_llave!ARM_STOCKN2 = 0
arm_llave!ARM_INGRESOS = 0
arm_llave!ARM_SALIDAS = 0
arm_llave!arm_cospro = 0
arm_llave!arm_stock2 = 0
arm_llave!ARM_COSTO_ULT = 0
arm_llave.Update

End Sub


Public Function ARMA_NOMBRE() As String
ARMA_NOMBRE = Trim(txtnombre.Text)
If LK_EMP = "CAM" Then
  If Val(Right(art_familia.Text, 6)) = 1 Then
  ''& " " & Trim(Left(art_marca.Text, 10)) & "-" & Trim(Left(art_numero.Text, 40)) & " " & Trim(Left(art_linea.Text, 5))
    ARMA_NOMBRE = Left(art_familia.Text, 1) & "." & Trim(Left(art_grupo.Text, 10)) & " " & Trim(Left(art_subfam.Text, 15))
    If Trim(Left(art_marca.Text, 10)) <> "" Then
      ARMA_NOMBRE = Trim(ARMA_NOMBRE) & " " & Trim(Left(art_marca.Text, 10))
    End If
    If Trim(Left(art_plancha.Text, 40)) <> "" Then
       ARMA_NOMBRE = Trim(ARMA_NOMBRE) & " " & Trim(Left(art_plancha.Text, 40))
    End If
    If Trim(Left(art_linea.Text, 5)) <> "" Then
        ARMA_NOMBRE = Trim(ARMA_NOMBRE) & " " & Trim(Left(art_linea.Text, 5))
    End If
    ARMA_NOMBRE = UCase(ARMA_NOMBRE)
End If
If Val(Right(art_familia.Text, 6)) = 4 Then
    ARMA_NOMBRE = Trim(Left(art_numero.Text, 15))
    If Trim(Left(art_grupo.Text, 15)) <> "" Then
      ARMA_NOMBRE = ARMA_NOMBRE & " " & Trim(Left(art_grupo.Text, 15))
    End If
    ARMA_NOMBRE = ARMA_NOMBRE & " " & Trim(Left(art_marca.Text, 15))
    ARMA_NOMBRE = UCase(ARMA_NOMBRE)
End If
  If Val(Right(art_familia.Text, 6)) = 2 Then
    'ARMA_NOMBRE = Trim(Left(art_numero.Text, 15)) & " " & Trim(Left(art_grupo.Text, 15))
    ARMA_NOMBRE = Trim(Left(art_numero.Text, 15)) & " " & Trim(Left(art_grupo.Text, 10)) '& " " & Trim(Left(art_subfam.Text, 15))
    If Trim(Left(art_marca.Text, 10)) <> "" Then
      ARMA_NOMBRE = Trim(ARMA_NOMBRE) & " " & Trim(Left(art_marca.Text, 10))
    End If
    If Trim(Left(art_plancha.Text, 10)) <> "" Then
      ARMA_NOMBRE = Trim(ARMA_NOMBRE) & " " & Trim(Left(art_plancha.Text, 10))
    End If
    If Trim(Left(art_linea.Text, 5)) <> "" Then
        ARMA_NOMBRE = Trim(ARMA_NOMBRE) & " " & Trim(Left(art_linea.Text, 5))
    End If
    ARMA_NOMBRE = UCase(ARMA_NOMBRE)
    
  
    ARMA_NOMBRE = UCase(ARMA_NOMBRE)
  End If
  Exit Function
End If

End Function

Public Sub PROD_PROC()

gridp.Clear
gridp.Cols = 3
gridp.Rows = 1
gridp.TextMatrix(0, 0) = "Descripción"
gridp.TextMatrix(0, 1) = "Codigo"
gridp.TextMatrix(0, 2) = "Productos"

gridp.ColWidth(0) = 1800
gridp.ColWidth(1) = 900
gridp.ColWidth(2) = 2500


PUB_TIPREG = 122
PUB_CODCIA = LK_CODCIA
If LK_EMP_PTO = "A" Then
   PUB_CODCIA = "00"
End If
SQ_OPER = 2
LEER_TAB_LLAVE
fila = 0
Do Until tab_mayor.EOF
 gridp.Rows = gridp.Rows + 1
 gridp.RowHeight(fila) = 285
 gridp.TextMatrix(fila + 1, 0) = Left(tab_mayor!tab_NOMLARGO, 40)
 fila = fila + 1
 If fila = 3 Then
   Exit Do
 End If
 tab_mayor.MoveNext
Loop
WCODART2 = Val(artloc_llave!ART_CODART2)
fila = 4
Do Until WCODART2 = 0
PSART_RELA.rdoParameters(0) = LK_CODCIA
PSART_RELA.rdoParameters(1) = WCODART2
art_rela.Requery
If art_rela.EOF Then
 MsgBox "Verificar la relacion ."
 Exit Sub
End If
fila = fila - 1
gridp.TextMatrix(fila, 1) = art_rela!art_alterno
gridp.TextMatrix(fila, 2) = art_rela!art_nombre
WCODART2 = Val(art_rela!ART_CODART2)
Loop

End Sub

Public Sub opcional()
Dim WWF As String
Dim ST_ACTUAL As Currency
Dim WfART_llave As rdoResultset
Dim WPSART_LLAVE As rdoQuery
Dim wcanti_unid As Currency
Dim WS_FILA As Integer
Dim xl  As Object
Dim E As Integer
Stop
Stop
Stop


pub_cadena = "select pa_codpa, sum(pa_prom) as tot  from paquetes where pa_codpa <> 0 and pa_codcia = ? group by pa_codpa"
Set PSART_RELA = CN.CreateQuery("", pub_cadena)
PSART_RELA(0) = 0
Set art_rela = PSART_RELA.OpenResultset(rdOpenKeyset, rdConcurValues)
PSART_RELA(0) = LK_CODCIA
art_rela.Requery

Do Until art_rela.EOF
SQ_OPER = 1
pu_codcia = LK_CODCIA
If art_rela!pa_codpa = 22242 Then Stop
PUB_CODART = art_rela!pa_codpa
LEER_ARM_LLAVE
If arm_llave.EOF Then
  MsgBox "no hay relacion  =  " & art_rela!pa_codpa
Else
  If Val(Nulo_Valor0(art_rela!TOT)) <> 0 Then
     arm_llave.Edit
     arm_llave!arm_saldo_s2 = Val(Nulo_Valor0(art_rela!TOT))
     arm_llave.Update
   End If
End If

art_rela.MoveNext
Loop


MsgBox " termino "
Exit Sub
Do Until art_rela.EOF
  SQ_OPER = 1
  PUB_TIPREG = 333
  PUB_NUMTAB = Val(art_rela!CLI_SUBGRUPO)
  PUB_CODCIA = LK_CODCIA
  LEER_TAB_LLAVE
  If tab_llave.EOF Then
      MsgBox "NO AMARRA CODIGO :  " & art_rela!cli_codclie
   '    art_rela.Edit
   '    art_rela!CLI_SUBGRUPO = 7
   '    art_rela!CLI_CASA1 = "07"
   '    art_rela.Update
'     GoTo salER
  End If
  If art_rela!CLI_MONEDA = "D" Then
      MsgBox "dOLARES  OJO " & art_rela!cli_codclie
      Stop
      GoTo salER
  End If
  
  If UCase(Left(Trim(art_rela!CLI_NOMBRE), 3)) = "AMB" Then
     art_rela.Edit
     art_rela!CLI_CASA1 = "06"
     art_rela!CLI_SUBGRUPO = 6
     art_rela.Update
  'Else
  '   art_rela!CLI_CASA1 = "07"
  End If
  
  
salER:
art_rela.MoveNext
Loop
MsgBox "acabo"
Exit Sub


Stop


SQ_OPER = 2
PUB_KEY = 0
pu_codcia = LK_CODCIA
LEER_ART_LLAVE
Do Until art_mayor.EOF
 SQ_OPER = 2
 pu_codcia = LK_CODCIA
 PUB_CODART = art_mayor!ART_KEY
 LEER_PRE_LLAVE
 If pre_mayor.RowCount > 1 Then
 If Trim(pre_mayor!PRE_FLAG_UNIDAD) <> "" Then
   pre_mayor.Edit
   pre_mayor!PRE_FLAG_UNIDAD = " "
   pre_mayor.Update
  
   pre_mayor.MoveNext
   pre_mayor.Edit
   pre_mayor!PRE_FLAG_UNIDAD = "A"
   pre_mayor.Update
  End If
 End If
art_mayor.MoveNext
Loop

Exit Sub
' RECALCULO DE LOS STOCK PARA TOMA DE INVENTARIO
MsgBox "LISTO PARA EMPEZAR"
'If xl Is Nothing Then
'   Set xl = CreateObject("Excel.Application")
'End If
DoEvents
pub_cadena = "SELECT DISTINCT FAR_NUMGUIA FROM FACART WHERE FAR_CODCIA = ? AND FAR_SERGUIA = ?  AND FAR_TIPMOV = 10 AND FAR_ESTADO <> 'E' ORDER BY  FAR_NUMGUIA DESC"
pub_cadena = "SELECT DISTINCT FAR_FBG, FAR_NUMSER, FAR_NUMFAC FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODART = ?  AND FAR_EQUIV = 1 AND  FAR_TIPMOV = 10 AND FAR_ESTADO <> 'E'"
Set WPSART_LLAVE = CN.CreateQuery("", pub_cadena)
WPSART_LLAVE(0) = 0
WPSART_LLAVE(1) = 0
Set WfART_llave = WPSART_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
WPSART_LLAVE(0) = LK_CODCIA


PUB_KEY = 0
pu_codcia = LK_CODCIA
SQ_OPER = 2
LEER_ART_LLAVE
WWF = ""
WS_FILA = 2
Do Until art_mayor.EOF

SQ_OPER = 2
PUB_CODART = art_mayor!ART_KEY
LEER_PRE_LLAVE
If pre_mayor.RowCount > 1 Then
 WPSART_LLAVE(1) = pre_mayor!PRE_codart
 WfART_llave.Requery
 If Not WfART_llave.EOF Then
   Debug.Print WfART_llave!far_fbg & "/" & WfART_llave!far_numser & " " & WfART_llave!far_numfac
 End If
 
End If


Ava00:
 
 art_mayor.MoveNext
Loop

MsgBox "LISTO"


Exit Sub


WS_FILA = 2
Do Until WS_FILA = 65000
 If Trim(xl.Cells(WS_FILA, 1)) = "" Then Exit Do
 SQ_OPER = 3
 pu_alterno = xl.Cells(WS_FILA, 1)
 pu_codcia = LK_CODCIA
 LEER_ART_LLAVE
 If art_llave_alt.EOF Then
   MsgBox "Notar codigo No Existe ...." & pu_alterno & " " & WS_FILA
   GoTo Ava001
 End If
 SQ_OPER = 1
 PUB_CODART = art_llave_alt!ART_KEY
 pu_codcia = LK_CODCIA
 LEER_ARM_LLAVE
 If arm_llave.EOF Then
   MsgBox " FALLO "
 End If
 WPSART_LLAVE(2) = arm_llave!ARM_CODART
 WfART_llave.Requery
 If WfART_llave.EOF Then
  ' MsgBox " NO SE AGREGO EN SALDO INCIAL "
   xl.Cells(WS_FILA, 8) = "MANUAL"
   GoTo Ava001
 End If
 SQ_OPER = 2
 pu_codcia = LK_CODCIA
 PUB_CODART = arm_llave!ARM_CODART
 LEER_PRE_LLAVE
 Do Until pre_mayor.EOF
    If Trim(pre_mayor!PRE_FLAG_UNIDAD) = "A" Then Exit Do
    
 pre_mayor.MoveNext
 Loop
 
 
 ST_ACTUAL = Format(pre_mayor!PRE_EQUIV * Val(xl.Cells(WS_FILA, 5)), "0.0000")
 If ST_ACTUAL <> Val(arm_llave!ARM_STOCK) Then
   WDIF = ST_ACTUAL - Val(arm_llave!ARM_STOCK)
   WfART_llave.Edit
   WfART_llave!far_cantidad_P = WfART_llave!far_cantidad
   WfART_llave!far_cantidad = WfART_llave!far_cantidad + WDIF
   WfART_llave.Update
   arm_llave.Edit
   arm_llave!arm_stock2 = arm_llave!ARM_STOCK
   arm_llave!ARM_STOCK = ST_ACTUAL 'Format(Val(xl.Cells(WS_FILA, 2)) * Val(xl.Cells(WS_FILA, 3)), "0.0000")
   
   arm_llave.Update
 End If
 
 
 
 
Ava001:
 WS_FILA = WS_FILA + 1
Loop

MsgBox "TERMINO PROCESO"
' TERMINO DE PROCESO
Exit Sub

' ACTUALIZA PRECIOS
Stop
pub_cadena = "SELECT * FROM ARTI , PRECIOS WHERE (ART_CODCIA = PRE_CODCIA) AND (ART_KEY = PRE_CODART) AND ART_FAMILIA = 1 AND PRE_SECUENCIA = 1 AND ART_CODCIA = ? "
Set WPSART_LLAVE = CN.CreateQuery("", pub_cadena)
WPSART_LLAVE(0) = LK_CODCIA
Set WfART_llave = WPSART_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
WPSART_LLAVE(0) = LK_CODCIA
WfART_llave.Requery

pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_SECUENCIA = 0 AND PRE_CODCIA = ? AND PRE_CODART = ? "
Set PSART_KEY = CN.CreateQuery("", pub_cadena)
PSART_KEY.rdoParameters(0) = "  "
PSART_KEY.rdoParameters(1) = 0
Set artloc_key = PSART_KEY.OpenResultset(rdOpenKeyset, rdConcurValues)
PSART_KEY.rdoParameters(0) = LK_CODCIA
Do Until WfART_llave.EOF
   PSART_KEY.rdoParameters(1) = WfART_llave!PRE_codart
   artloc_key.Requery
   artloc_key.Edit
   artloc_key!PRE_PRE1 = Format(WfART_llave!PRE_PRE1 / WfART_llave!PRE_EQUIV, "0.0000")
   artloc_key!PRE_LITRO = Format(Nulo_Valor0(WfART_llave!PRE_LITRO) / WfART_llave!PRE_EQUIV, "0.000")
   artloc_key!pre_PESO = Format(WfART_llave!pre_PESO / WfART_llave!PRE_EQUIV, "0.00")
   artloc_key!PRE_FLAG_UNIDAD = " "
   artloc_key.Update
   WfART_llave.Edit
   WfART_llave!PRE_FLAG_UNIDAD = "A"
   WfART_llave.Update
   WfART_llave.MoveNext
Loop
MsgBox "FIN"
Exit Sub

' FIN DE ACTUALIZAR




Stop
Stop
Stop
If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If
DoEvents
' lblproceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
DoEvents

WPAS = "131296"
xl.Workbooks.Open "C:\CARGA\ALMACEN.xls"  ', 0, True, 4,  WPAS, WPAS


xl.Application.Visible = True
WS_FILA = 3
Stop
Do Until WS_FILA = 65000
If Trim(xl.Cells(WS_FILA, 1)) = "" Then Exit Do
If Val(xl.Cells(WS_FILA, 25)) = 0 Then GoTo Ava0022
SQ_OPER = 3
pu_alterno = xl.Cells(WS_FILA, 1)
pu_codcia = LK_CODCIA
LEER_ART_LLAVE
If art_llave_alt.EOF Then
  MsgBox "Notar codigo No Existe ...." & pu_alterno
  GoTo Ava0022
End If
SQ_OPER = 2
pu_codcia = LK_CODCIA
PUB_CODART = art_llave_alt!ART_KEY
LEER_PRE_LLAVE
Do Until pre_mayor.EOF
  If pre_mayor!PRE_FLAG_UNIDAD = "A" Then
     Exit Do
  End If
  pre_mayor.MoveNext
Loop
pre_mayor.Edit

pre_mayor!PRE_PRE1 = Format(Val(xl.Cells(WS_FILA, 16)) * (1 + (LK_IGV / 100)), "0.00")
pre_mayor!PRE_LITRO = Val(xl.Cells(WS_FILA, 8))
pre_mayor!pre_PESO = Val(xl.Cells(WS_FILA, 9))
pre_mayor.Update

Ava0022:
WS_FILA = WS_FILA + 1

Loop


MsgBox "PARE STOP"
MsgBox "PARE STOP"
MsgBox "PARE STOP"
MsgBox "PARE STOP"
MsgBox "PARE STOP"
MsgBox "PARE STOP"
Exit Sub
Stop

WS_FILA = 2
'Do Until dataO.Recordset.EOF
PUB_NUMSER = 1
PUB_NUMFAC = 1
WS_NUMSEC = 0
Do Until WS_FILA = 65000
If Trim(xl.Cells(WS_FILA, 1)) = "" Then Exit Do
If Val(xl.Cells(WS_FILA, 7)) = 0 Then GoTo Ava
If Val(xl.Cells(WS_FILA, 25)) = 0 Then GoTo Ava
wcanti_unid = Format(Val(xl.Cells(WS_FILA, 7)) * Val(xl.Cells(WS_FILA, 25)), "0.0000")
  
SQ_OPER = 3
pu_alterno = xl.Cells(WS_FILA, 1)
pu_codcia = LK_CODCIA
LEER_ART_LLAVE
If art_llave_alt.EOF Then
  MsgBox "Notar codigo No Existe ...." & pu_alterno
  GoTo Ava
End If
SQ_OPER = 1
PUB_CODART = art_llave_alt!ART_KEY
pu_codcia = LK_CODCIA
LEER_ARM_LLAVE
arm_llave.Edit
arm_llave!ARM_STOCK = Abs(wcanti_unid)
arm_llave.Update
arm_llave.Requery

far_llave.AddNew
      far_llave!FAR_TIPMOV = 6
      far_llave!FAR_CODCIA = LK_CODCIA
      far_llave!far_cod_sunat = 0 'Val(Right(i_codsunat.Text, 5))
      far_llave!far_numser = PUB_NUMSER
      far_llave!FAR_CODVEN = 0
      far_llave!far_numfac = PUB_NUMFAC + 1
      WS_NUMSEC = WS_NUMSEC + 1
      far_llave!FAR_NUMSEC = WS_NUMSEC
      far_llave!FAR_STOCK = Abs(wcanti_unid)
      far_llave!far_codart = Val(arm_llave!ARM_CODART)
      far_llave!far_cantidad = Abs(wcanti_unid)
      far_llave!FAR_PRECIO = Format(Val(xl.Cells(WS_FILA, 14)) / Val(xl.Cells(WS_FILA, 7)), "0.0000")
      far_llave!FAR_equiv = 1 'Val(xl.Cells(WS_FILA, 7))
      far_llave!far_descri = "PZA" 'Trim(xl.Cells(WS_FILA, 6))
      far_llave!far_PESO = 0
      far_llave!far_signo_car = 0
      far_llave!far_signo_car = 0
      If wcanti_unid < 0 Then
        far_llave!far_signo_arm = -1
      Else
        far_llave!far_signo_arm = 1
      End If
      far_llave!far_codclie = 0
      far_llave!FAR_MONEDA = "S"
      far_llave!FAR_EX_IGV = 0
      far_llave!FAR_cp = " "
      far_llave!FAR_fecha_compra = LK_FECHA_DIA
      far_llave!far_estado = "N"
      far_llave!FAR_estado2 = "N"
      far_llave!FAR_COSPRO = 0
      far_llave!FAR_COSPRO_ANT = 0
      
      far_llave!far_fbg = " "
      far_llave!far_IMPTO = 0
      far_llave!FAR_TOT_FLETE = 0
      far_llave!FAR_FLETE = 0
      far_llave!FAR_DESCTO = 0
      far_llave!FAR_TOT_DESCTO = 0
      far_llave!FAR_GASTOS = 0
      far_llave!FAR_BRUTO = 0
      far_llave!FAR_NUMDOC = 1
      far_llave!far_numguia = 0
      far_llave!far_serguia = 0
      far_llave!FAR_pordescto1 = 0
      far_llave!FAR_costeo = "A"
      far_llave!FAR_COSTEO_REAL = "A"
      far_llave!FAR_tipo_cambio = 1
      far_llave!FAR_DIAS = 0
      far_llave!FAR_fecha = LK_FECHA_DIA
      far_llave!FAR_NUMSER_C = 0
      far_llave!FAR_NUMFAC_C = 1
      far_llave!FAR_NUMOPER = 1
      far_llave!far_precio_neto = 0
'      far_llave!FAR_CONSIG = 0
      far_llave!far_otra_cia = " "
      far_llave!far_subtra = 1
      far_llave!far_transito = " "
      far_llave!far_subtra = " "
      far_llave!far_otra_cia = " "
      far_llave!far_transito = " "
      far_llave!far_subtra = " "
      far_llave!far_JABAS = 0
      far_llave!far_UNIDADES = 0
      far_llave!far_mortal = 0
      far_llave!far_num_precio = 0
      far_llave!FAR_ORDEN_UNIDADES = 0
      far_llave!FAR_SUBTOTAL = 0
      'far_llave!far_ISLA = 0
      far_llave!far_turno = 0
      far_llave!far_concepto = " "
      far_llave!far_concepto = "Saldo Inicial de Inventario"
      far_llave!far_codusu = LK_CODUSU
      far_llave!FAR_HORA = Format(Now, "hh:mm:ss AMPM")
      far_llave!FAR_NUM_LOTE = 0
      far_llave!FAR_PEDSER = 0
      far_llave!FAR_PEDFAC = 0
      far_llave!far_pedsec = 0
      'far_llave!far_fbg2 = 0
      far_llave!FAR_TIPDOC = "IN"
      far_llave.Update
Ava:
WS_FILA = WS_FILA + 1
Loop

MsgBox "TERMNO"
Exit Sub




'Dim BDRUTA
'BDRUTA = "E:\SALDO\bd2.mdb"
'dataO.DatabaseName = BDRUTA
'dataO.RecordSource = "STOCKCOL"
'dataO.Refresh
'If dataO.Recordset.EOF Then
'  MsgBox "VERIFICAR ...NO HAY DATOS EN dataO..", 48, Pub_Titulo
'  Exit Sub
'End If


'pub_cadena = "SELECT * FROM COMAEST WHERE COM_CODCIA = '55'  ORDER BY COM_CUENTA"
pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ? "
Set WPSART_LLAVE = CN.CreateQuery("", pub_cadena)
Set WfART_llave = WPSART_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
WPSART_LLAVE(0) = LK_CODCIA
WfART_llave.Requery
' AGREGAR POR MESES !!!!!!!!!!!!!!!
Stop
BDRUTA = "C:\ST.mdb"
PUB_FECHA = CDate("30/06/2000")
dataO.DatabaseName = BDRUTA
dataO.RecordSource = "STOCKCOL"
dataO.Refresh
If dataO.Recordset.EOF Then
  MsgBox "VERIFICAR ...NO HAY DATOS EN dataO..", 48, Pub_Titulo
  Exit Sub
End If
barra.Visible = True
barra.Min = 0
barra.max = dataO.Recordset.RecordCount
barra.Value = 0
ww_conta = 0
flag_otro = ""
WS_NRO_MOV = -1
'WCVOUCHER = Val(Mid(Trim(DataO.Recordset!NUMVOU), 3, Len(Trim(DataO.Recordset!NUMVOU))))
Stop
WC_TIPMOV = -1
Do Until dataO.Recordset.EOF
    DoEvents
    If UCase(Trim(dataO.Recordset!DETMOV)) = "APERTURA" Then GoTo PASA_APERTURA
    ww_conta = ww_conta + 1
PASA_APERTURA:
dataO.Recordset.MoveNext
Loop

Screen.MousePointer = 0
MsgBox "TERMINO ..... SIGUE CHEQUEO DE CUENTAS"

Exit Sub
Exit Sub


'PARA AGREGAR ARTICULOS
'' *************************
If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If
DoEvents
'lblproceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
DoEvents
WPAS = "131296"
xl.Workbooks.Open "C:\CARGA\ALMACEN.xls"  ', 0, True, 4,  WPAS, WPAS
xl.Application.Visible = True
WS_FILA = 2

Do Until Trim(xl.Cells(WS_FILA, 1)) = ""
    If Trim(xl.Cells(WS_FILA, 7)) = 0 Then
      MsgBox "no agregado"
      GoTo SALTA_ARTI_001
    End If
    
SALTA_ARTI_001:
    WS_FILA = WS_FILA + 1
Loop
MsgBox " TEWRMINO "

Exit Sub

Stop
WS_FILA = 156
xl.Application.Visible = True


Exit Sub

'PARA AGREGAR ARTICULOS
'' *************************
If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If
DoEvents
'lblproceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
DoEvents
WPAS = "131296"
xl.Workbooks.Open "C:\CARGA\ALMACEN.xls"  ', 0, True, 4,  WPAS, WPAS
xl.Application.Visible = True
WS_FILA = 2

Do Until Trim(xl.Cells(WS_FILA, 1)) = ""
    If Trim(xl.Cells(WS_FILA, 7)) = 0 Then
    MsgBox "no agregado"
'    GoTo SALTA_ARTI
    End If
    pu_alterno = Trim(xl.Cells(WS_FILA, 1))
    pu_codcia = LK_CODCIA
    art_llave_alt.Requery
    


    'art_subfam.ListIndex = 0
    DS.ListIndex = 1
    txtnombre.Text = Trim(xl.Cells(WS_FILA, 2)) '+ " " + Trim(xl.Cells(ws_fila, 4)) 'Trim(dataO.Recordset!DESCRIP)
    txt_alterno.Text = Trim(xl.Cells(WS_FILA, 1))
    'If Left(Trim(xl.Cells(WS_FILA, 1)), 1) = "P" Then
   '   cheservi(1).Value = True
   ' Else
   '   cheservi(0).Value = True
   ' End If
 '  cmddolares_Click
   If Trim(xl.Cells(WS_FILA, 7)) > 1 Then
     grid_UNID_KeyDown 45, 0
     grid_unid.TextMatrix(2, 0) = Trim(xl.Cells(WS_FILA, 6))
     grid_unid.TextMatrix(2, 1) = Trim(xl.Cells(WS_FILA, 7))
     grid_unid.TextMatrix(2, 5) = Format((Val(xl.Cells(WS_FILA, 16)) * (1 + (LK_IGV / 100))), "0.00")
   '   grid_unid.TextMatrix(2, 21) = Trim(xl.Cells(WS_FILA, 11)) ' SOLES
     grid_unid.TextMatrix(2, 16) = Format((Val(xl.Cells(WS_FILA, 16)) * (1 + (LK_IGV / 100))), "0.00") ' DOLARES
      grid_unid.TextMatrix(2, 28) = Trim(xl.Cells(WS_FILA, 8)) ' CANTIDAD DE LITROS
      grid_unid.TextMatrix(2, 26) = Trim(xl.Cells(WS_FILA, 9)) ' CANTIDAD DE PESO
      ' UNIDAD MINIMA
     grid_unid.TextMatrix(1, 5) = Format(Val(grid_unid.TextMatrix(2, 5)) / Val(grid_unid.TextMatrix(2, 1)), "0.00")
     ' grid_unid.TextMatrix(1, 21) = Format(Val(grid_unid.TextMatrix(2, 5)) / Val(grid_unid.TextMatrix(2, 1)), "0.00")
     grid_unid.TextMatrix(1, 16) = Format(Val(grid_unid.TextMatrix(2, 5)) / Val(grid_unid.TextMatrix(2, 1)), "0.00")
     grid_unid.TextMatrix(1, 28) = Format(Val(grid_unid.TextMatrix(2, 28)) / Val(grid_unid.TextMatrix(2, 1)), "0.00")
     grid_unid.TextMatrix(1, 26) = Format(Val(grid_unid.TextMatrix(2, 26)) / Val(grid_unid.TextMatrix(2, 1)), "0.00")
    
    Else
     grid_unid.TextMatrix(1, 0) = Trim(xl.Cells(WS_FILA, 6))
     grid_unid.TextMatrix(1, 5) = Trim(xl.Cells(WS_FILA, 11))
     'grid_unid.TextMatrix(1, 21) = Trim(xl.Cells(WS_FILA, 11)) ' SOLES
     grid_unid.TextMatrix(1, 16) = Trim(xl.Cells(WS_FILA, 16)) ' dolares
     grid_unid.TextMatrix(1, 28) = Trim(xl.Cells(WS_FILA, 8)) ' CANTIDAD DE LITROS
     grid_unid.TextMatrix(1, 26) = Trim(xl.Cells(WS_FILA, 9)) ' CANTIDAD DE PESO

    End If
    
    cmdagregar_Click
SALTA_ART33I:
    WS_FILA = WS_FILA + 1
Loop
MsgBox " TEWRMINO "




Exit Sub
Stop
If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If
DoEvents
' lblproceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
DoEvents


xl.Workbooks.Open "C:\CARGA\ZONAS.xls"  ', 0, True, 4,  WPAS, WPAS

xl.Application.Visible = True
WS_FILA = 2
Stop
Do Until WS_FILA = 65000
If Trim(xl.Cells(WS_FILA, 1)) = "" Then Exit Do
  tab_llave.Requery
  tab_llave.AddNew
  tab_llave!TAB_CODCIA = LK_CODCIA
  tab_llave!TAB_TIPREG = 35
  tab_llave!TAB_NUMTAB = Val(Trim(xl.Cells(WS_FILA, 1)))
  tab_llave!tab_NOMLARGO = Trim(xl.Cells(WS_FILA, 2))
  tab_llave!tab_nomcorto = Trim(xl.Cells(WS_FILA, 2))
  tab_llave!TAB_CODART = 0
  tab_llave!TAB_CODCLIE = 0
  tab_llaveTAB_CONTABLE2 = 0

  tab_llave.Update



WS_FILA = WS_FILA + 1

Loop


MsgBox " TERMINO"

End Sub

Private Sub cmd_AddItem_Click()
Dim Cnn_SQL As New ADODB.Connection
Dim Cnn_DBF As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim RSPRE As New ADODB.Recordset
Dim rs_T As New ADODB.Recordset
Dim RSEquivalencia As New ADODB.Recordset
Dim sCnnDBF As String
Dim sCnnSQL As String
Dim i As Long
Dim s_Sql As String
Dim s_Fam As String
Dim s_Marca As String
Dim i_NumTab As Integer
Dim RES As Integer
Dim sCosto As Double
Dim iCount As Integer
Dim Icount2 As Integer
Dim NumReg As Integer
Dim equivalencia As Integer
Dim RSTabNomlargo As New ADODB.Recordset

On Error GoTo ErrHandle

If Trim(LK_CODUSU) <> "ADMIN" Then
    MsgBox "!!!! ..... Acceso denegado para este tipo de procesos. .... !!!!! Consulte al Administrador", vbExclamation, "Import data"
    GoTo SALIR
End If

    RES = MsgBox("Desea Ingresar Stock o Articulos", vbYesNo, Pub_Titulo)
    If RES = vbYes Then
        GoTo Inventario
    End If
    
    RES = MsgBox("!!!! ..... Este proceso inserta articulos nuevos al sistema. Esta seguro de realizar este proceso.", vbExclamation + vbYesNo, "Import data")
    If RES = vbNo Then GoTo SALIR
   '' sCnnDBF = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=DSN_DBF" '"Provider=SQLOLEDB.1;Persist Security Info=False;pwd=;User ID=sa;Initial Catalog=BDHUEM;Data Source=PC01" '"Provider=MSDASQL;Data Source=DSN_TMP"
    
  ''  Cnn_DBF.CursorLocation = adUseClient
   '' Cnn_DBF.Open sCnnDBF
    
    sCnnSQL = "Provider=SQLOLEDB.1;Persist Security Info=False;pwd=anteromariano;User ID=sa;Initial Catalog=BDATOS;Data Source=laptop" '"Provider=MSDASQL;Data Source=DSN_DATOS" '"provider=SQLOLEDB;Data source=PC01;initial catalog=bdatos;password=;user id=sa"
    Cnn_SQL.CursorLocation = adUseClient
    Cnn_SQL.Open sCnnSQL
 
    's_Sql = "select * from artidiro"
    s_Sql = "select * from cargart where cod_art in (select cod_art from artidiro) "
    
    rs.Open s_Sql, Cnn_SQL, adOpenForwardOnly, adLockOptimistic, adCmdText
 
    pgb_Progress.Visible = True
    pgb_Progress.max = rs.RecordCount + 1
    NumReg = rs.RecordCount
    While Not rs.EOF
        DoEvents
        iCount = iCount + 1
        lblprogress.Caption = CStr(iCount) & " de " & CStr(NumReg)
        pgb_Progress.Value = iCount
        If IsNull(rs("cod_art")) Then GoTo SEGUIR
        If Len(rs("cod_art")) < 2 Then GoTo SEGUIR
        
        If Left(cmdAgregar.Caption, 2) = "&G" Then
            cmdCancelar = True
        End If
        
'        SQ_OPER = 3
'        pu_alterno = Trim(rs("codi_item"))
'        pu_codcia = LK_CODCIA
'        LEER_ART_LLAVE
'        If Not art_llave_alt.EOF Then
'            GoTo SEGUIR
'        End If
        
        cmdAgregar = True
        txt_alterno = rs("cod_art")
        
        txtnombre = Trim(rs("des_art")) & " " & Trim(rs("grupito"))
        txtnombre = Replace(txtnombre, "'", "´")
        'txtMin.Text = Nulo_Valor0(rs("MIN_ART"))
        'txtMax.Text = Nulo_Valor0(rs("max_art"))
        
'                F14.Visible = False
        grid_unid.Rows = 2
        If IsNull(rs("moneda")) Then
            LK_MONEDA = "S"
        Else
            If rs("moneda") = 1 Then
                LK_MONEDA = "S"
            Else
                LK_MONEDA = "D"
            End If
        End If
        
        DS.Text = LK_MONEDA
        
        grid_unid.TextMatrix(1, 14) = "A" 'flag
        'sCosto = Format(Nulo_Valor0(Trim(rs("valo_cost"))), "0.0000")
        'grid_unid.TextMatrix(1, 3) = sCosto
        'grid_unid.TextMatrix(1, 27) = sCosto
        
        'grid_unid.TextMatrix(1, 5) = Format(Nulo_Valor0(rs("precio_v")), "0.0000")
        equivalencia = 1
        If IsNull(rs("uni_med")) Then
            UNIDAD = "UNIDAD"
        Else
            UNIDAD = Trim(UCase(rs("uni_med")))
        End If
        
        grid_unid.TextMatrix(1, 1) = equivalencia
        grid_unid.TextMatrix(1, 0) = UNIDAD
        'PRECIOS SOLES
        If LK_MONEDA = "S" Then
            grid_unid.TextMatrix(1, 21) = Format(Nulo_Valor0(rs("precio_V")), "0.00")
            grid_unid.TextMatrix(1, 22) = Format(Nulo_Valor0(rs("precio_2")), "0.00")
            grid_unid.TextMatrix(1, 23) = Format(Nulo_Valor0(rs("precio_3")), "0.00")
            grid_unid.TextMatrix(1, 24) = Format(Nulo_Valor0(rs("precio_4")), "0.0000")
            'grid_unid.TextMatrix(1, 25) = Format(Nulo_Valor0(rs("precio_5")), "0.0000")
            'grid_unid.TextMatrix(1, 26) = Format(Nulo_Valor0(rs("precio_v")), "0.0000")
            'grid_unid.TextMatrix(1, 27) = Format(Nulo_Valor0(rs("precio_6")), "0.0000")
        Else
            'PRECIOS DOLARES
            grid_unid.TextMatrix(1, 16) = Format(Nulo_Valor0(rs("precio_V")), "0.00")
            grid_unid.TextMatrix(1, 17) = Format(Nulo_Valor0(rs("precio_2")), "0.00")
            grid_unid.TextMatrix(1, 18) = Format(Nulo_Valor0(rs("precio_3")), "0.00")
            grid_unid.TextMatrix(1, 19) = Format(Nulo_Valor0(rs("precio_4")), "0.00")
            'grid_unid.TextMatrix(1, 20) = Format(Nulo_Valor0(rs("precio_5")), "0.0000")
            'grid_unid.TextMatrix(1, 21) = Format(Nulo_Valor0(rs("precio_v")), "0.0000")
            'grid_unid.TextMatrix(1, 22) = Format(Nulo_Valor0(rs("precio_6")), "0.0000")
        End If
        
        
        'para la segunda unidad
'        If IsNull(rs("mayo_unid")) Then
'            GoTo CONTINUAR
'        Else
'            If UCase(Trim(rs("mayo_unid"))) = UCase(Trim(rs("defe_unid"))) Then GoTo CONTINUAR
'            UNIDAD = Nulo_Valors(rs("mayo_unid"))
'            equivalencia = rs("dequ_unid")
'            If equivalencia = 0 Then equivalencia = 1
'        End If
'        grid_unid.Rows = 3
'        grid_unid.TextMatrix(2, 0) = UNIDAD
'        grid_unid.TextMatrix(2, 5) = Format(Nulo_Valor0(rs("valo_vent")) * equivalencia, "0.0000")
'        'PRECIOS SOLES
'        If LK_MONEDA = "S" Then
'            grid_unid.TextMatrix(2, 21) = Format(Nulo_Valor0(rs("valo_vent")) * equivalencia, "0.0000")
'        Else
'            'PRECIOS DOLARES
'            grid_unid.TextMatrix(2, 16) = Format(Nulo_Valor0(rs("valo_vent") * equivalencia), "0.0000")
'        End If
'        grid_unid.TextMatrix(2, 1) = equivalencia

CONTINUAR:

    '================FAMILIA=======================
        RSTabNomlargo.Open "SELECT GRUPITO FROM ARTIDIRO WHERE RUBRO='" & rs("RUBRO") & "'", Cnn_SQL, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not RSTabNomlargo.EOF Then
            s_Fam = Left(UCase(Trim(Nulo_Valors(RSTabNomlargo("grupito")))), 40) ''DESCRIPCION DE FAMILIA
        Else
            s_Fam = ""
        End If
        RSTabNomlargo.Close
        
      's_Fam = Left(UCase(Trim(Nulo_Valors(rs("grupito")))), 40) ''DESCRIPCION DE FAMILIA
      If s_Fam = "" Then
        art_familia.ListIndex = 0
      Else
        If FindInCmb(art_familia, s_Fam) Then
        Else
            s_Sql = "select max(tab_numtab) as max_n from tablas where tab_codcia='" & LK_CODCIA & "' and tab_tipreg=122"
            rs_T.Open s_Sql, Cnn_SQL, adOpenForwardOnly, adLockOptimistic, adCmdText
            i_NumTab = IIf(IsNull(rs_T!max_n), 0, rs_T!max_n) + 1
            WTMP = i_NumTab
            Set rs_T = Nothing
            s_Sql = "insert into tablas (tab_codcia,tab_tipreg,tab_numtab,tab_nomlargo,TAB_NOMCORTO)"
            s_Sql = s_Sql + " VALUES('" & LK_CODCIA & "','122','" & i_NumTab & "','" & Left(s_Fam, 40) & "','" & Mid(s_Fam, 1, 10) & "')"
            Cnn_SQL.Execute s_Sql
            LLENADO_FAM
            If FindInCmb(art_familia, s_Fam) Then
            End If
        End If
      End If
    art_familia_LostFocus
    '================DIVISION=======================
      WTMP = Val(Right(art_familia.Text, 6))
      
        RSTabNomlargo.Open "SELECT SUBGRUPITO FROM ARTIDIRO WHERE SRUBRO='" & rs("SRUBRO") & "'", Cnn_SQL, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not RSTabNomlargo.EOF Then
            s_Marca = Left(UCase(Trim(Nulo_Valors(RSTabNomlargo("SUBgrupito")))), 40) ''DESCRIPCION DE FAMILIA
        Else
            s_Marca = ""
        End If
        RSTabNomlargo.Close
        
        
      's_Marca = Left(UCase(Trim(Nulo_Valors(rs("subgrupito")))), 40) 'DESCRIPCION
      If s_Marca = "" Then
        s_Marca = "OTRA"
      Else
        s_Marca = Replace(s_Marca, "'", " ")
      End If
      If s_Marca = "" Then
       art_subfam.ListIndex = 0
      Else
        If FindInCmb(art_subfam, s_Marca) Then
        Else
            s_Sql = "select max(tab_numtab) as max_n from tablas where tab_codcia='" & LK_CODCIA & "' and tab_tipreg=123"
            rs_T.Open s_Sql, Cnn_SQL, adOpenForwardOnly, adLockOptimistic, adCmdText
            i_NumTab = IIf(IsNull(rs_T!max_n), 0, rs_T!max_n) + 1
            Set rs_T = Nothing
            s_Sql = "insert into tablas (tab_codcia,tab_tipreg,tab_numtab,tab_nomlargo,TAB_CODART,TAB_NOMCORTO)"
            s_Sql = s_Sql + " VALUES('" & LK_CODCIA & "','123','" & i_NumTab & "','" & Left(s_Marca, 40) & "'," & WTMP & ",'" & Mid(s_Marca, 1, 10) & "')"
            Cnn_SQL.Execute s_Sql
            LLENADO_SUBFAM art_subfam, WTMP
            If FindInCmb(art_subfam, s_Marca) Then
            End If
        End If
      End If
      art_subfam_LostFocus
    '================================
  
  
    '================LINEA=======================
'      WTMP = Val(Right(art_subfam.Text, 6))
'      s_Linea = UCase(Trim(Nulo_Valors(rs("DESC_ARTI"))))  'DESCRIPCION
'      If s_Linea = "" Then
'        s_Linea = "OTRA"
'      Else
'        s_Linea = Replace(s_Linea, "'", " ")
'      End If
'      If s_Linea = "" Then
'       art_grupo.ListIndex = 0
'      Else
'        If FindInCmb(art_grupo, s_Linea) Then
'        Else
'            s_Sql = "select max(tab_numtab) as max_n from tablas where tab_codcia='" & LK_CODCIA & "' and tab_tipreg=129"
'            rs_T.Open s_Sql, Cnn_SQL, adOpenForwardOnly, adLockOptimistic, adCmdText
'            i_NumTab = IIf(IsNull(rs_T!max_n), 0, rs_T!max_n) + 1
'            Set rs_T = Nothing
'            s_Sql = "insert into tablas (tab_codcia,tab_tipreg,tab_numtab,tab_nomlargo,TAB_CODART,TAB_NOMCORTO)"
'            s_Sql = s_Sql + " VALUES('" & LK_CODCIA & "','129','" & i_NumTab & "','" & s_Linea & "'," & WTMP & ",'" & Mid(s_Linea, 1, 10) & "')"
'            Cnn_SQL.Execute s_Sql
'            LLENADO_SUBFAM art_grupo, WTMP
'            If FindInCmb(art_grupo, s_Linea) Then
'            End If
'        End If
'      End If
'    art_grupo_LostFocus
    '================SUBLINEA=======================
'    WTMP = Val(Right(art_grupo.Text, 6))
'      s_SubLinea = Nulo_Valors(Trim(rs("desc_subline")))  'DESCRIPCION
'      If s_SubLinea = "" Then
'        s_SubLinea = "OTRA"
'      Else
'        s_SubLinea = Replace(s_SubLinea, "'", " ")
'      End If
'      If s_SubLinea = "" Then
'       art_numero.ListIndex = 0
'      Else
'        If FindInCmb(art_numero, s_SubLinea) Then
'        Else
'            s_Sql = "select max(tab_numtab) as max_n from tablas where tab_codcia='" & LK_CODCIA & "' and tab_tipreg=130"
'            rs_T.Open s_Sql, Cnn_SQL, adOpenForwardOnly, adLockOptimistic, adCmdText
'            i_NumTab = IIf(IsNull(rs_T!max_n), 0, rs_T!max_n) + 1
'            Set rs_T = Nothing
'            s_Sql = "insert into tablas (tab_codcia,tab_tipreg,tab_numtab,tab_nomlargo,TAB_CODART,TAB_NOMCORTO)"
'            s_Sql = s_Sql + " VALUES('" & LK_CODCIA & "','130','" & i_NumTab & "','" & s_SubLinea & "'," & WTMP & ",'" & Mid(s_SubLinea, 1, 10) & "')"
'            Cnn_SQL.Execute s_Sql
'            LLENADO_SUBFAM art_numero, WTMP
'            If FindInCmb(art_numero, s_SubLinea) Then
'            End If
'        End If
'      End If
    '================================
' MARCA
            
      's_Linea = UCase(Trim(Nulo_Valors(rs("des_art")))) ''DESCRIPCION DE MARCA
     ' If Right(s_Linea, 1) <> " " Then
     '  s_Linea = ""
     '   GoTo otra_marca
     ' End If
     ' iPos = InStrRev(s_Linea, "(")
     ' s_Linea = Right(s_Linea, Len(s_Linea) - iPos)
     ' s_Linea = Left(s_Linea, Len(s_Linea) - 1)
'otra_marca:
 '     If s_Linea = "" Then
  '      art_linea.ListIndex = 0
   '   Else
    '    If FindInCmb(art_linea, s_Linea) Then
     '   Else
      '      s_Sql = "select max(tab_numtab) as max_n from tablas where tab_codcia='" & LK_CODCIA & "' and tab_tipreg=131"
       '     rs_T.Open s_Sql, Cnn_SQL, adOpenForwardOnly, adLockOptimistic, adCmdText
        '    i_NumTab = IIf(IsNull(rs_T!max_n), 0, rs_T!max_n) + 1
         '   WTMP = i_NumTab
          '  Set rs_T = Nothing
           ' s_Sql = "insert into tablas (tab_codcia,tab_tipreg,tab_numtab,tab_nomlargo,TAB_NOMCORTO)"
         '   s_Sql = s_Sql + " VALUES('" & LK_CODCIA & "','131','" & i_NumTab & "','" & Left(s_Linea, 40) & "','" & Mid(s_Linea, 1, 10) & "')"
         '   Cnn_SQL.Execute s_Sql
         '   LLENADO_LINEA 131
         '   If FindInCmb(art_linea, s_Linea) Then
         '   End If
       ' End If
     ' End If
   ' art_Linea_LostFocus
    
  cmdAgregar = True

SEGUIR:

  DoEvents
  If Not rs.AbsolutePosition = adPosEOF Then
   pgb_Progress.Value = rs.AbsolutePosition
  End If
  

  rs.MoveNext
 Wend
 rs.Close
 pgb_Progress.Visible = Not True
 lblprogress.Caption = ""
 MsgBox "Proceso terminado satisfactoriamente", vbInformation, "Import Arti"
 GoTo SALIR
 
Exit Sub

'=====================================================================================================================
'=====================================================================================================================
'=====================================================================================================================
Inventario:
Dim RSStock As New ADODB.Recordset
Dim PRECIO As Double
Dim sTotal As Double
Dim RSOrigen As New ADODB.Recordset
Dim tempcodigo As String
Dim pcosto As Double
Dim moneda As Double

     's_Cnn = "Provider=MSDASQL;Data Source=DSN_DBF"
     'Cnn_DBF.CursorLocation = adUseClient
     'Cnn_DBF.Open s_Cnn
        sCnnSQL = "Provider=SQLOLEDB.1;Persist Security Info=False;pwd=anteromariano;User ID=sa;Initial Catalog=BDATOS;Data Source=laptop" '"Provider=MSDASQL;Data Source=DSN_DATOS" '"provider=SQLOLEDB;Data source=PC01;initial catalog=bdatos;password=;user id=sa"
        Cnn_SQL.CursorLocation = adUseClient
        Cnn_SQL.Open sCnnSQL
    
     s_Sql = "select * from cargart WHERE STK_ART<>0"
     's_Sql = "SELECT A.COD_ART,A.UNI_MED, A1.PRECIO_C,A.MONEDA,A.STK_ART " 'a.DES_ART,a.UNI_MED,a.GRUPITO,a.SUBGRUPITO,a.CAN_UNI,a.RUBRO,a.SRUBRO,a.MIN_ART,a.MAX_ART,a1.PRECIO_C,a.PRECIO_V,a.PORCIENTO,a.STK_ART,a.MONEDA,a.PRECIO_1,a.PRECIO_2,a.PRECIO_3,a.FECHA_A
     's'_Sql = s_Sql & " FROM cargart A INNER JOIN cargart1 A1 ON A.COD_ART = A1.COD_ART AND A.RUBRO = A1.RUBRO AND A.SRUBRO = A1.SRUBRO "
     
    
     
     RSOrigen.Open s_Sql, Cnn_SQL, adOpenForwardOnly, adLockOptimistic, adCmdText
     
     pgb_Progress.Visible = True
     pgb_Progress.max = RSOrigen.RecordCount + 1
     NumReg = RSOrigen.RecordCount
     
    PUB_NUMFAC = 1
    'RSORIGEN.Open "SELECT * FROM carga", ConnAcc

    Do While Not RSOrigen.EOF
        DoEvents
        iCount = iCount + 1
        lblprogress.Caption = CStr(iCount) & " de " & CStr(NumReg)
        pgb_Progress.Value = iCount
        
        If Trim(Nulo_Valors(RSOrigen("COD_ART"))) = "" Then GoTo Ava
        If Nulo_Valor0(RSOrigen("STK_ART")) = 0 Then GoTo Ava
        If tempcodigo = Trim(Nulo_Valors(RSOrigen("COD_ART"))) Then GoTo Ava
        
        s_Sql = "select * from ARTIDIRO WHERE COD_ART='" & RSOrigen("cod_art") & "'"
        RSStock.Open s_Sql, Cnn_SQL, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not RSStock.EOF Then
            'pcosto = Nulo_Valor0(RSStock("PRECIO_C"))
            moneda = IIf(IsNull(RSStock("MONEDA")), 1, RSStock("MONEDA"))
        Else
            pcosto = 0
        End If
        RSStock.Close
        
        'pcosto = Nulo_Valor0(RSOrigen("PRECIO_C"))
        
        SQ_OPER = 3
        pu_alterno = Trim(RSOrigen("COD_ART"))
        pu_codcia = LK_CODCIA
        LEER_ART_LLAVE
        If art_llave_alt.EOF Then
            MsgBox "Arti no existe=" & RSOrigen("COD_ART")
            GoTo Ava
        End If
'        If Left(Trim(art_llave_alt("art_nombre")), 50) <> Left(Trim(RSOrigen("DES_ART")), 50) Then
'            MsgBox "dupli"
'            GoTo Ava
'        End If
        wcanti_unid = Val(RSOrigen("STK_ART"))
        SQ_OPER = 1
        PUB_CODART = art_llave_alt!ART_KEY
        pu_codcia = LK_CODCIA
        LEER_ARM_LLAVE
        arm_llave.Edit
        arm_llave!ARM_STOCK = wcanti_unid
        arm_llave!ARM_INGRESOS = wcanti_unid
        arm_llave!ARM_Saldo_n = wcanti_unid
        arm_llave.Update
        arm_llave.Requery
        
        If WS_NUMSEC = 20 Then
            GoSub LlenarAllog
            PUB_NUMFAC = PUB_NUMFAC + 1
            sTotal = 0
            WS_NUMSEC = 0
        End If
        If moneda = 1 Then
            PRECIO = pcosto
        ElseIf moneda = 2 Then
            PRECIO = pcosto * 3.5 'verificar el tipo de cambio de acuerdo a la empresa
        End If
        PRECIO = Format(PRECIO / 1.19, "0.0000")
        sTotal = sTotal + PRECIO * wcanti_unid
        
        far_llave.AddNew
        far_llave!FAR_TIPMOV = 6
        far_llave!FAR_CODCIA = LK_CODCIA
        far_llave!far_cod_sunat = 0
        far_llave!far_numser = PUB_NUMSER
        far_llave!FAR_CODVEN = 0
        far_llave!far_numfac = PUB_NUMFAC
        WS_NUMSEC = WS_NUMSEC + 1
        far_llave!FAR_NUMSEC = WS_NUMSEC
        far_llave!FAR_STOCK = wcanti_unid
        far_llave!far_codart = Val(arm_llave!ARM_CODART)
        far_llave!far_cantidad = wcanti_unid
        far_llave!FAR_PRECIO = PRECIO
        far_llave!FAR_equiv = 1 'RSOrigen("MEQU_UNID")
        far_llave!far_descri = Trim(Nulo_Valors(RSOrigen("UNI_MED")))
        far_llave!far_PESO = 0
        far_llave!far_signo_car = 0
        far_llave!far_signo_car = 0
        far_llave!far_signo_arm = 1
        far_llave!far_key_dircli = 0
        far_llave!far_codclie = 0
        far_llave!FAR_MONEDA = "S"
        far_llave!FAR_EX_IGV = 0
        far_llave!FAR_cp = " "
        far_llave!FAR_fecha_compra = LK_FECHA_DIA
        far_llave!far_estado = "N"
        far_llave!FAR_estado2 = "N"
        far_llave!FAR_COSPRO = 0
        far_llave!FAR_COSPRO_ANT = 0
        far_llave!far_fbg = " "
        far_llave!far_IMPTO = 0
        far_llave!FAR_TOT_FLETE = 0
        far_llave!FAR_FLETE = 0
        far_llave!FAR_DESCTO = 0
        far_llave!FAR_TOT_DESCTO = 0
        far_llave!FAR_GASTOS = 0
        far_llave!FAR_BRUTO = 0
        far_llave!FAR_NUMDOC = 1
        far_llave!far_numguia = 0
        far_llave!far_serguia = 0
        far_llave!FAR_pordescto1 = 0
        far_llave!FAR_costeo = "A"
        far_llave!FAR_COSTEO_REAL = "A"
        far_llave!FAR_tipo_cambio = 1
        far_llave!FAR_DIAS = 0
        far_llave!FAR_fecha = LK_FECHA_DIA
        far_llave!FAR_NUMSER_C = 0
        far_llave!FAR_NUMFAC_C = 1
        far_llave!FAR_NUMOPER = 1
        far_llave!far_precio_neto = 0
        far_llave!far_otra_cia = " "
        far_llave!far_subtra = 1
        far_llave!far_transito = " "
        far_llave!far_subtra = " "
        far_llave!far_otra_cia = " "
        far_llave!far_transito = " "
        far_llave!far_subtra = " "
        far_llave!far_JABAS = 0
        far_llave!far_UNIDADES = 0
        far_llave!far_mortal = 0
        far_llave!far_num_precio = 0
        far_llave!FAR_ORDEN_UNIDADES = 0
        far_llave!FAR_SUBTOTAL = 0
        far_llave!far_turno = 0
        far_llave!far_concepto = "Saldo Inicial de Inventario"
        far_llave!far_codusu = LK_CODUSU
        far_llave!FAR_HORA = Format(Now, "hh:mm:ss AMPM")
        far_llave!FAR_NUM_LOTE = 0
        far_llave!FAR_PEDSER = 0
        far_llave!FAR_PEDFAC = 0
        far_llave!far_pedsec = PUB_VISITA
        far_llave!FAR_TIPDOC = "IN"
        far_llave.Update
Ava:
        WS_FILA = WS_FILA + 1
        GoTo sLoop
        
LlenarAllog:
        all_llave.AddNew
        all_llave!all_CODCIA = LK_CODCIA
        all_llave!ALL_FECHA_DIA = LK_FECHA_DIA
        all_llave!ALL_NUMOPER = PUB_NUMFAC
        all_llave!ALL_CODTRA = 2403
        all_llave!all_flag_ext = 0
        all_llave!ALL_CODCLIE = 0
        all_llave!ALL_CODART = 0
        all_llave!ALL_IMPORTE_AMORT = sTotal
        all_llave!ALL_IMPORTE = 0
        all_llave!ALL_CHESER = i_c
        all_llave!all_SECUENCIA = PUB_NUMFAC
        all_llave!ALL_IMPORTE_DOLL = 0
        all_llave!all_codusu = "ADMIN"
        all_llave!ALL_PRECIO = 0
        all_llave!ALL_CODVEN = 0
        all_llave!ALL_FBG = ""
        all_llave!ALL_CP = ""
        all_llave!ALL_TIPDOC = ""
        all_llave!ALL_CANTIDAD = 0
        all_llave!ALL_NUMGUIA = 0
        all_llave!all_codban = 0
        all_llave!ALL_autocon = "Carga Inicial:INVENTARIO INICIAL"
        all_llave!all_chenum = 0
        all_llave!ALL_CHESEC = 0
        all_llave!ALL_NUMSER = 0
        all_llave!all_numfac = PUB_NUMFAC
        all_llave!ALL_FECHA_VCTO = LK_FECHA_DIA
        all_llave!all_neto = sTotal
        all_llave!ALL_BRUTO = sTotal
        all_llave!ALL_GASTOS = 0
        all_llave!ALL_IMPTO = 0
        all_llave!ALL_DESCTO = 0
        all_llave!ALL_MONEDA_CAJA = "S"
        all_llave!ALL_moneda_ccm = ""
        all_llave!ALL_MONEDA_CLI = "S"
        all_llave!ALL_NUMDOC = 0
        all_llave!ALL_LIMCRE_ANT = 0
        all_llave!ALL_LIMCRE_ACT = 0
        'all_llave!all_CANT_CHEQ = 0
        all_llave!all_NUM_OPER2 = 0
        all_llave!all_sIGNO_ARM = 1
        all_llave!all_codtra_ext = 2403
        all_llave!ALL_SIGNO_CCM = 0
        all_llave!ALL_SIGNO_CAR = 0
        all_llave!ALL_SIGNO_CAJA = 0
        all_llave!ALL_tipmov = 6
        all_llave!all_numser_c = 0
        all_llave!all_numfac_c = 0
        all_llave!ALL_serdoc = 0
        all_llave!ALL_TIPO_CAMBIO = 3.5
        all_llave!ALL_flete = 0
        all_llave!ALL_SUBTRA = "Carga Inicial"
        all_llave!ALL_HORA = Time
        all_llave!all_FACART = ""
        all_llave!all_concepto = "Aumento de Inventario"
        all_llave!ALL_numoper2 = PUB_NUMFAC
        all_llave!all_FECHA_ANT = LK_FECHA_DIA
        all_llave!ALL_SITUACION = ""
        all_llave!ALL_FECHA_SUNAT = LK_FECHA_DIA
        all_llave!ALL_FLAG_SO = "X"
        all_llave!ALL_IMPG1 = 0
        all_llave!ALL_IMPG2 = 0
        all_llave!ALL_CTAG1 = ""
        all_llave!ALL_CTAG2 = ""
        all_llave!ALL_CODSUNAT = 2
        all_llave!ALL_FECHA_PRO = LK_FECHA_DIA
        all_llave!ALL_FECHA_CAN = LK_FECHA_DIA
        all_llave!all_SERIE_REC = 0
        all_llave!ALL_NUM_RECIBO = 0
        all_llave!ALL_RUC = ""
        all_llave.Update
        Return
sLoop:
    tempcodigo = Trim(Nulo_Valors(RSOrigen("COD_ART")))
    RSOrigen.MoveNext
Loop

MsgBox "TERMiNO"
Exit Sub

'=========================================================================================================
'=========================================================================================================
'=========================================================================================================


 
 
ErrHandle:
 MsgBox Err.Description, vbExclamation
SALIR:
 Set rs = Nothing
 Set RSPRE = Nothing
 ''Resume Next
 Set rs_T = Nothing
 Set Cnn_DBF = Nothing
 Set Cnn_SQL = Nothing
End Sub

Private Function FindInCmb(ByVal cbo As ComboBox, ByVal s_Familia As String) As Boolean
Dim i As Long
Dim aux As Boolean
Dim aux_f As String

    aux = False
    For i = 0 To cbo.ListCount - 1
     aux_f = cbo.List(i)
     aux_f = Trim$(Left$(aux_f, Len(aux_f) - 10))
     If Trim(aux_f) = Trim(s_Familia) Then
      cbo.ListIndex = i
      aux = True
      Exit For
     End If
    Next
    FindInCmb = aux
    
End Function

'MIC====================================================
'agregue para buscador de articulos segun estructura
    
Private Sub cmdconsultar_Click()
Dim sCodigo As String
Dim SQL As String
Dim sMoneda As String
Dim UNIDAD As String
Dim Familia As Integer
Dim SubFam As Integer
Dim grupo As Integer
Dim Linea As Integer
On Error GoTo Handler
    
    
    Familia = Val(Right(artfamilia.Text, 6))
    SubFam = Val(Right(artsubfam.Text, 6))
    grupo = Val(Right(artgrupo.Text, 6))
    Linea = Val(Right(artlinea.Text, 6))
    
       
    SQL = "SELECT ARTI.ART_key, ARTI.ART_NOMBRE AS Articulo, ARTI.ART_ALTERNO AS Alterno, PRECIOS.PRE_UNIDAD, ARTICULO.ARM_STOCK "
    SQL = SQL & "FROM ARTI INNER JOIN PRECIOS ON ARTI.ART_KEY = PRECIOS.PRE_CODART AND ARTI.ART_CODCIA = PRECIOS.PRE_CODCIA "
    SQL = SQL & " INNER JOIN ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA "
    SQL = SQL & "WHERE ARTI.ART_CODCIA = '" & LK_CODCIA & "' AND PRECIOS.PRE_FLAG_UNIDAD='A'"
    SQL = SQL & " AND ART_FAMILIA = " & Familia
    SQL = SQL & " AND ART_SUBFAM = " & SubFam
    SQL = SQL & " AND ART_SUBGRU = " & grupo
    If Linea >= 0 Then
    SQL = SQL & " AND ART_Linea = " & Linea
    End If
    Set RDQPRECIOS = CN.CreateQuery("", SQL)
    Set RDRPRECIOS = RDQPRECIOS.OpenResultset(rdOpenKeyset, rdConcurValues)
    RDRPRECIOS.Requery
    
    grdarticulos.Clear
    SETGRID
    grdarticulos.Rows = IIf(RDRPRECIOS.RowCount = 0, 2, RDRPRECIOS.RowCount + 1)
    
    fila = 0
    Do While Not RDRPRECIOS.EOF
        fila = fila + 1
        'grdarticulos.RowHeight(fila) = 285
        grdarticulos.Row = fila
        grdarticulos.TextMatrix(fila, 1) = Trim(RDRPRECIOS!Alterno)
        grdarticulos.TextMatrix(fila, 2) = Trim(RDRPRECIOS!articulo)
        grdarticulos.TextMatrix(fila, 3) = Trim(RDRPRECIOS!pre_unidad)
        grdarticulos.TextMatrix(fila, 4) = RDRPRECIOS!ARM_STOCK
        grdarticulos.TextMatrix(fila, 5) = RDRPRECIOS!ART_KEY
        If (fila \ 2) * 2 = fila Then BackColorRow fila
        RDRPRECIOS.MoveNext
    Loop
    grdarticulos.SetFocus
    grdarticulos.COL = 1
    grdarticulos.Row = 1
    Exit Sub
Handler:
    MsgBox Err.Description, vbCritical, Pub_Titulo
End Sub
Private Sub artfamilia_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        frmbusqueda.Visible = False
    End If
End Sub

Private Sub artfamilia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        artsubfam.SetFocus
        DoEvents
        SendKeys "%{UP}"
        Exit Sub
    End If
End Sub
Private Sub artfamilia_LostFocus()
Dim wpos As Integer
Dim WFAMI2 As Integer
    If Trim(artfamilia.Text) = "" Then
        artsubfam.Clear
        Exit Sub
    End If
    wpos = artsubfam.ListIndex
    WFAMI2 = Val(Trim(Right(artfamilia.Text, 6)))
    PUB_TIPREG = 123
    LLENADO_SUBFAM artsubfam, WFAMI2
    On Error GoTo sigue
    artsubfam.ListIndex = wpos
    Exit Sub
sigue:
    Resume Next
End Sub
Private Sub artsubfam_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        frmbusqueda.Visible = False
    End If
End Sub
Private Sub artsubfam_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        artgrupo.SetFocus
        DoEvents
        SendKeys "%{UP}"
        Exit Sub
    End If
End Sub
Private Sub artsubfam_LostFocus()
Dim wpos As Integer
Dim WFAMI2 As Integer
    If Trim(artsubfam.Text) = "" Then
        artgrupo.Clear
        Exit Sub
    End If
    wpos = artgrupo.ListIndex
    WFAMI2 = Val(Trim(Right(artsubfam.Text, 6)))
    PUB_TIPREG = 129
    LLENADO_SUBFAM artgrupo, WFAMI2
    On Error GoTo sigue
    artgrupo.ListIndex = wpos
    Exit Sub
sigue:
    Resume Next
End Sub
Private Sub artgrupo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        frmbusqueda.Visible = False
    End If
End Sub
Private Sub artgrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        artlinea.SetFocus
        DoEvents
        SendKeys "%{UP}"
        Exit Sub
    End If
End Sub
Private Sub artlinea_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        frmbusqueda.Visible = False
    End If
End Sub
Private Sub artlinea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grdarticulos.SetFocus
    End If
End Sub
Private Sub artlinea_LostFocus()
    cmdconsultar_Click
End Sub
Private Sub grdarticulos_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        frmbusqueda.Visible = False
        If Txt_key.Enabled Then Txt_key.SetFocus
        If txt_alterno.Enabled Then txt_alterno.SetFocus
    End If
End Sub
Private Sub grdarticulos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If grdarticulos.TextMatrix(grdarticulos.Row, 5) <> "" Then
            If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
                txt_alterno.Text = Trim(grdarticulos.TextMatrix(grdarticulos.Row, 1))
                pu_alterno = Trim(txt_alterno.Text)
                txt_alterno_KeyPress 13
            Else
                Txt_key.Text = Trim(grdarticulos.TextMatrix(grdarticulos.Row, 5))
                PUB_KEY = Val(Txt_key.Text)
                txt_key_KeyPress 13
            End If
            frmbusqueda.Visible = False
        End If
    End If
End Sub
Private Sub SETGRID()
    grdarticulos.Clear
    grdarticulos.FormatString = "|Codigo|Descripcion|Unidad|Stock"
    grdarticulos.Rows = 2
    grdarticulos.Cols = 6
    grdarticulos.ColWidth(0) = 0
    grdarticulos.ColWidth(1) = 1500 '|Codigo
    grdarticulos.ColWidth(2) = 7000 '|Descricion
    grdarticulos.ColWidth(3) = 1200 '|Unidad
    grdarticulos.ColWidth(4) = 1000 '|Stock
    grdarticulos.ColWidth(5) = 0   '|art_key
End Sub
Private Sub BackColorRow(ByVal iRow As Long)
Dim iCol As Long
    grdarticulos.Row = iRow
    For iCol = 1 To grdarticulos.Cols - 1
     grdarticulos.COL = iCol
     grdarticulos.CellBackColor = &H80000018
     'grid_unid.CellFontBold = True
    Next
End Sub


