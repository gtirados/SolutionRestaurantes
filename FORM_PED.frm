VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FORM_PED 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Ordenes de Compra"
   ClientHeight    =   4890
   ClientLeft      =   1500
   ClientTop       =   1140
   ClientWidth     =   6600
   ControlBox      =   0   'False
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
   Icon            =   "FORM_PED.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4890
   ScaleWidth      =   6600
   Tag             =   "55"
   WindowState     =   2  'Maximized
   Begin ComctlLib.ListView ListView1 
      Height          =   495
      Left            =   8685
      TabIndex        =   31
      Top             =   7305
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   10235904
      BackColor       =   16118252
      BorderStyle     =   1
      Appearance      =   1
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
   Begin VB.Frame fraListar 
      Caption         =   "Seleccionar"
      ForeColor       =   &H00800000&
      Height          =   6780
      Left            =   4215
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   7650
      Begin VB.CommandButton cmdSelecAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Desmarcar Todos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1755
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   6195
         Width           =   1530
      End
      Begin VB.CommandButton cmdSelecAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Seleccionar Todos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   6195
         Width           =   1530
      End
      Begin VB.ListBox fami 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   7410
      End
      Begin VB.OptionButton opcomo 
         Caption         =   "x División"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   4200
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton opcomo 
         Caption         =   "x Sub División"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   5880
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton opcomo 
         Caption         =   "x Proveedor"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   2220
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton opcomo 
         Caption         =   "x Articulo"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.ListBox lisarti 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   13
         Top             =   3180
         Width           =   7425
      End
      Begin VB.CommandButton cmdcancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Retornar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   6720
         Picture         =   "FORM_PED.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6030
         Width           =   795
      End
      Begin VB.CommandButton cmdpasar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Procesar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   5865
         Picture         =   "FORM_PED.frx":0C10
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   6030
         Width           =   795
      End
      Begin VB.ListBox subfami 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Left            =   3555
         TabIndex        =   7
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox i_codart2 
         Height          =   315
         Left            =   1005
         MaxLength       =   16
         TabIndex        =   30
         Top             =   495
         Width           =   1740
      End
      Begin VB.Label lfami 
         Caption         =   "División:"
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
         Left            =   165
         TabIndex        =   12
         Tag             =   "9999"
         Top             =   480
         Width           =   885
      End
      Begin VB.Label lsubfami 
         AutoSize        =   -1  'True
         Caption         =   "Sub División:"
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
         Left            =   3570
         TabIndex        =   11
         Tag             =   "9999"
         Top             =   495
         Width           =   1065
      End
      Begin VB.Label larti 
         AutoSize        =   -1  'True
         Caption         =   "Articulos"
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
         Left            =   135
         TabIndex        =   8
         Tag             =   "9999"
         Top             =   2955
         Width           =   750
      End
      Begin VB.Label lblarti 
         Caption         =   "Articulo :"
         ForeColor       =   &H009C3000&
         Height          =   255
         Left            =   150
         TabIndex        =   32
         Top             =   525
         Width           =   840
      End
   End
   Begin VB.CommandButton MODIF 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Modificar"
      Height          =   690
      Left            =   10650
      MaskColor       =   &H00F5F1EC&
      Picture         =   "FORM_PED.frx":1A22
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   2865
      Width           =   1110
   End
   Begin VB.CommandButton cmdIngreso 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Ingreso"
      Height          =   690
      Left            =   10650
      MaskColor       =   &H00F5F1EC&
      Picture         =   "FORM_PED.frx":28BC
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   3555
      Width           =   1110
   End
   Begin VB.CommandButton SALIR 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ce&rrar"
      Height          =   690
      Left            =   10650
      MaskColor       =   &H00F5F1EC&
      Picture         =   "FORM_PED.frx":33A6
      Style           =   1  'Graphical
      TabIndex        =   66
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   5610
      Width           =   1110
   End
   Begin VB.CommandButton cancelar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancelar"
      Height          =   690
      Left            =   10650
      MaskColor       =   &H00F5F1EC&
      Picture         =   "FORM_PED.frx":3C1C
      Style           =   1  'Graphical
      TabIndex        =   65
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   4905
      Width           =   1110
   End
   Begin VB.PictureBox MSChart1 
      Height          =   975
      Left            =   480
      ScaleHeight     =   915
      ScaleWidth      =   6075
      TabIndex        =   39
      Top             =   7440
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Frame f1 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   2055
      Left            =   75
      TabIndex        =   18
      Top             =   0
      Width           =   11775
      Begin VB.CommandButton cmdMostrar 
         BackColor       =   &H00F5F1EC&
         Caption         =   " Articulos"
         Height          =   615
         Left            =   10455
         Picture         =   "FORM_PED.frx":43CA
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   960
         Width           =   1080
      End
      Begin VB.ListBox liscia 
         Height          =   735
         Left            =   7560
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   62
         Top             =   180
         Width           =   3855
      End
      Begin VB.ComboBox moneda 
         BackColor       =   &H00F5F1EC&
         Height          =   315
         ItemData        =   "FORM_PED.frx":4C2C
         Left            =   6000
         List            =   "FORM_PED.frx":4C36
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   990
         Width           =   1215
      End
      Begin VB.CheckBox chesin 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Orden Sin Valorizar"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   7560
         TabIndex        =   53
         Top             =   1020
         Width           =   1980
      End
      Begin VB.CheckBox checia 
         BackColor       =   &H00C0C0C0&
         Caption         =   "02 - Calcular Sucursal"
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   5160
         TabIndex        =   52
         Top             =   120
         Width           =   2220
      End
      Begin VB.CheckBox cheneg 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Quitar los Negativos"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   7560
         TabIndex        =   51
         Top             =   1320
         Width           =   2340
      End
      Begin VB.ComboBox fpago 
         BackColor       =   &H00F5F1EC&
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   1635
         Width           =   2052
      End
      Begin VB.ComboBox txtagencia 
         BackColor       =   &H00F5F1EC&
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   1020
         Width           =   3135
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FORM_PED.frx":4C4E
         Left            =   120
         List            =   "FORM_PED.frx":4C50
         Style           =   1  'Simple Combo
         TabIndex        =   43
         Top             =   1635
         Width           =   2415
      End
      Begin VB.TextBox txtcontacto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2880
         TabIndex        =   41
         Top             =   1635
         Width           =   2175
      End
      Begin VB.ComboBox codpro 
         BackColor       =   &H00F5F1EC&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   360
         Width           =   4935
      End
      Begin VB.TextBox txtdoc 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   675
         TabIndex        =   22
         Top             =   1020
         Width           =   975
      End
      Begin VB.TextBox tserie 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   21
         Top             =   1005
         Width           =   495
      End
      Begin VB.TextBox txtdias 
         Height          =   300
         Left            =   8880
         MaxLength       =   3
         TabIndex        =   20
         Text            =   "7"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox tpromedio 
         Height          =   300
         Left            =   10770
         MaxLength       =   3
         TabIndex        =   19
         Text            =   "90"
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Compañia:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   7560
         TabIndex        =   63
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   255
         Left            =   5175
         TabIndex        =   59
         Top             =   1005
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   255
         Left            =   5160
         TabIndex        =   57
         Top             =   480
         Width           =   735
      End
      Begin VB.Label fecha 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   6000
         TabIndex        =   56
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lcodart 
         BackStyle       =   0  'Transparent
         Caption         =   "F/Pago"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   255
         Index           =   7
         Left            =   5160
         TabIndex        =   45
         Tag             =   "9999"
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label lcodart 
         BackStyle       =   0  'Transparent
         Caption         =   "Enviar a:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   44
         Tag             =   "9999"
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lcodart 
         BackStyle       =   0  'Transparent
         Caption         =   "Att."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   42
         Tag             =   "9999"
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   135
         Width           =   1215
      End
      Begin VB.Label lcodart 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº. Doc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   27
         Tag             =   "9999"
         Top             =   735
         Width           =   885
      End
      Begin VB.Label lcodart 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad Dias a Solicitar :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   375
         Index           =   1
         Left            =   7485
         TabIndex        =   26
         Tag             =   "9999"
         Top             =   1605
         Width           =   1365
      End
      Begin VB.Label lcodart 
         BackStyle       =   0  'Transparent
         Caption         =   "Dias Promedio :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   375
         Index           =   2
         Left            =   9720
         TabIndex        =   25
         Tag             =   "9999"
         Top             =   1605
         Width           =   1125
      End
      Begin VB.Label lcodart 
         BackStyle       =   0  'Transparent
         Caption         =   "Agencia:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   24
         Tag             =   "9999"
         Top             =   765
         Width           =   765
      End
   End
   Begin ComctlLib.ProgressBar Barra 
      Height          =   315
      Left            =   3720
      TabIndex        =   2
      Tag             =   "0"
      Top             =   3480
      Visible         =   0   'False
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   0
      Min             =   77
      Max             =   91
   End
   Begin VB.Frame ESTADO 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Articulos :"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   4695
      Left            =   60
      TabIndex        =   1
      Tag             =   "100"
      Top             =   2085
      Width           =   11805
      Begin VB.TextBox txtdescto 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F1EC&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3951
         TabIndex        =   71
         Top             =   4305
         Width           =   1335
      End
      Begin VB.TextBox textovar 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F1EC&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2745
         TabIndex        =   64
         Top             =   1830
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CommandButton Imprimir 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Im&primir"
         Height          =   690
         Left            =   10590
         Picture         =   "FORM_PED.frx":4C52
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   2130
         Width           =   1110
      End
      Begin VB.TextBox imax 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   10710
         TabIndex        =   48
         Text            =   "22"
         Top             =   390
         Width           =   810
      End
      Begin VB.ComboBox UNIDAD 
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
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txttotal 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F1EC&
         Enabled         =   0   'False
         Height          =   315
         Left            =   9120
         TabIndex        =   35
         Top             =   4305
         Width           =   1335
      End
      Begin VB.TextBox txtigv 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F1EC&
         Enabled         =   0   'False
         Height          =   315
         Left            =   6535
         TabIndex        =   34
         Top             =   4305
         Width           =   1335
      End
      Begin VB.TextBox txtvalorv 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F1EC&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1367
         TabIndex        =   33
         Top             =   4320
         Width           =   1335
      End
      Begin ComctlLib.ProgressBar PB 
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2640
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid grid_fac 
         Height          =   4020
         Left            =   120
         TabIndex        =   0
         Tag             =   "9999"
         Top             =   255
         Width           =   10350
         _ExtentX        =   18256
         _ExtentY        =   7091
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         BackColor       =   16118252
         BackColorFixed  =   4210752
         ForeColorFixed  =   16777215
         BackColorBkg    =   16777215
         GridColor       =   8421504
         GridColorFixed  =   12632256
         FocusRect       =   2
         HighLight       =   2
         GridLinesFixed  =   1
         AllowUserResizing=   3
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
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descto:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   2734
         TabIndex        =   72
         Tag             =   "9999"
         Top             =   4320
         Width           =   1185
      End
      Begin VB.Label lblpeso 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F1EC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1320
         TabIndex        =   55
         Top             =   4680
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Peso(Kg.):"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   195
         TabIndex        =   54
         Top             =   4665
         Width           =   1095
      End
      Begin VB.Label lblmoneda 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   8760
         TabIndex        =   50
         Top             =   4350
         Width           =   300
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Max. de Item :"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   10575
         TabIndex        =   49
         Top             =   165
         Width           =   1110
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   3
         Left            =   7902
         TabIndex        =   38
         Tag             =   "9999"
         Top             =   4305
         Width           =   1185
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "I.G.V. :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   2
         Left            =   5318
         TabIndex        =   37
         Tag             =   "9999"
         Top             =   4320
         Width           =   1185
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   150
         TabIndex        =   36
         Tag             =   "9999"
         Top             =   4320
         Width           =   1185
      End
      Begin VB.Label momen 
         Caption         =   "Un Momento ..."
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   1560
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   120
      Top             =   7440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblbarraos 
      Alignment       =   2  'Center
      BackColor       =   &H008B4914&
      Caption         =   "Solution"
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
      Left            =   -120
      TabIndex        =   60
      Top             =   6840
      Width           =   11895
   End
End
Attribute VB_Name = "FORM_PED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xl As Object
Dim PSPRO_V As rdoQuery
Dim PRO_V As rdoResultset
Dim LOC_CODCIA As String * 2
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset
Dim PS_REP02 As rdoQuery
Dim llave_rep02 As rdoResultset
Dim PS_REP03 As rdoQuery
Dim llave_rep03 As rdoResultset
Dim PS_REP04 As rdoQuery
Dim llave_rep04 As rdoResultset

Dim VAR_ACTIVAR As Integer
Dim WCOD_ORIGINAL As Currency
Dim WPASA As Boolean
Dim WSELE As String * 1
Dim llave1
Dim loc_key
Dim fila As Integer
Dim ws_bruto_d, ws_bruto_h As Currency
Dim SUM_D As Currency
Dim SUM_H As Currency
Dim PSTEMP_LLAVE As rdoQuery
Dim temp_llave As rdoResultset
Dim WMODO As String * 1
Dim LOC_ITEM As Integer
Dim cop_llave As rdoResultset
Dim PSCOP_LLAVE As rdoQuery
Dim LOC_CANCELA As Integer
Dim PSTEMP_MAYOR As rdoQuery
Dim temp_mayor As rdoResultset

Dim temporal
Dim wfila_act As Integer
Dim loc_ini As Integer
Dim loc_fin  As Integer
Dim Wsec As Integer


Dim PSLOC_WARTI As rdoQuery
Dim llave_sum_arti   As rdoResultset
Dim FLAG_MODI As Integer
Option Explicit

Private Sub cancelar_Click()
WMODO = ""
cmdIngreso.Caption = "&Ingreso"
f1.Enabled = False
ESTADO.Enabled = False
PB.Visible = False
barra.Visible = False
fila = 0
SUM_D = 0
SUM_H = 0
LIMPIA_DATOS
CABE_MAN
FORM_PED.MODIF.Enabled = True
fecha.Caption = ""

'grid_fac.SetFocus
FORM_PED.MSChart1.Visible = False
End Sub

Private Sub cancelar_GotFocus()
fraListar.Visible = False
End Sub

Private Sub cmdcancel_Click()
On Error GoTo SALE

fraListar.Visible = False
If grid_fac.Enabled Then grid_fac.SetFocus
Exit Sub
SALE:
End Sub

Private Sub cmdIngreso_Click()
Dim flag_precio As String * 1
Dim ws_tot_debe, ws_tot_haber As Currency
Dim er As rdoError
Dim pub_mensaje As String
Const ingre = 2
Const MODIF = 1
Dim N As Integer
Dim LOC_SALDO_CAR As Currency
Dim FLAG As Boolean
Dim pub_mensaje_err As String
Dim WS_NRO_MOV, ws_nro_voucher As Long
Dim w_dh  As String
If Left(cmdIngreso.Caption, 2) = "&G" Then

FLAG = False
ws_tot_debe = Val(grid_fac.TextMatrix(1, 4))
ws_tot_haber = Val(grid_fac.TextMatrix(1, 5))

'If ws_tot_debe = 0 And ws_tot_haber = 0 Then
'  MsgBox "Ingrese los Vouchers ..", 48, Pub_Titulo
'  If grid_fac.Rows > 2 Then
'     grid_fac.Col = 1
'     grid_fac.Row = 2
'     grid_fac.SetFocus
'   End If
  barra.Visible = False
'  GoTo fin
'End If
If grid_fac.Rows <= 2 Then
  MsgBox "Ingresar datos ...", 48, Pub_Titulo
  cmdMostrar.SetFocus
  GoTo fin
End If
If Trim(codpro.Text) = "" Then
  MsgBox "Ingresar Proveedor ...", 48, Pub_Titulo
  codpro.SetFocus
  GoTo fin
End If
If Trim(txtagencia.Text) = "" Then
  MsgBox "Ingresar Agencia de Transporte ...", 48, Pub_Titulo
  codpro.SetFocus
  GoTo fin
End If
suma_grid
flag_precio = ""
For fila = 2 To grid_fac.Rows - 1
 If grid_fac.TextMatrix(fila, 1) <> "" Then
  If Val(grid_fac.TextMatrix(fila, 2)) <= 0 Then
    MsgBox "Verificar, cantidad en cero o menor. - " & grid_fac.TextMatrix(fila, 1) & " : " & grid_fac.TextMatrix(fila, 0), 48, Pub_Titulo
    grid_fac.SetFocus
    GoTo fin
  End If
  If Val(grid_fac.TextMatrix(fila, 4)) = 0 Then
    flag_precio = "A"
    ''MsgBox "Verificar hay algun precio en 0 .", 48, Pub_Titulo
    ''grid_fac.SetFocus
 '   GoTo fin
  End If
End If
Next fila
If flag_precio = "A" Then
  pub_mensaje = "Existen precios com valor Cero.   ¿Desea Continuar... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then
     GoTo fin
  End If
End If
Screen.MousePointer = 11
DoEvents
barra.Visible = True
DoEvents
barra.Min = 0
barra.max = fila
barra.Value = 0
exito = True
barra.Value = 1
If FLAG_MODI = 1 Then
   temp_llave.MoveFirst
   Do Until temp_llave.EOF
      temp_llave.Edit
      temp_llave.Delete
      temp_llave.MoveNext
   Loop
End If

GoSub ACT1
pub_mensaje = "Usted. desea imprimir la Orden de Compra... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbYes Then
  Imprimir_Click
End If
fila = 1
SUM_D = 0
SUM_H = 0
CABE_MAN
LIMPIA_DATOS
fila = 0
If Pub_Respuesta <> vbYes Then
cancelar.SetFocus
End If
barra.Visible = False
cmdIngreso.Caption = "&Ingreso"

GoTo fin

ACT1:

fila = 1
ws_nro_voucher = ws_nro_voucher + 1
FLAG = False
WS_NRO_MOV = 0
fila = 2
Do While FLAG = False
   If Trim(grid_fac.TextMatrix(fila, 1)) = "" Then GoTo pasa
    ' grabo todo
   temp_llave.AddNew
   temp_llave!PED_CODCIA = LK_CODCIA
   temp_llave!PED_FECHA = LK_FECHA_DIA
   temp_llave!PED_NUMSER = Trim(tserie.Text)
   temp_llave!PED_NUMFAC = Val(txtdoc.Text)
   temp_llave!PED_NUMSEC = WS_NRO_MOV
   temp_llave!PED_CANTIDAD = Val(grid_fac.TextMatrix(fila, 2))
   temp_llave!PED_PRECIO = Val(grid_fac.TextMatrix(fila, 4))
   temp_llave!PED_CODUSU = LK_CODUSU
   temp_llave!PED_IGV = Val(txtigv.Text)
   temp_llave!PED_BRUTO = Val(txtvalorv.Text)
   temp_llave!PED_CONTACTO = Nulo_Valors(txtContacto.Text)
   temp_llave!PED_ESTADO = "N"
   temp_llave!PED_CODUSU = LK_CODUSU
   temp_llave!PED_CODART = Val(grid_fac.TextMatrix(fila, 9))
   temp_llave!PED_UNIDAD = Trim(grid_fac.TextMatrix(fila, 3))
   temp_llave!PED_OFERTA = txtagencia.Text
   temp_llave!PED_NOMCLIE = Combo1.Text
   temp_llave!PED_FORMA = fpago.Text
   temp_llave!PED_RUCCLIE = txtagencia.ListIndex
   temp_llave!PED_MONEDA = Left(moneda.Text, 1)
   temp_llave!PED_SUBTOTAL = Trim(grid_fac.TextMatrix(fila, 5))
   temp_llave!PED_EQUIV = 1
   temp_llave!PED_CODCLIE = Val(Right(codpro.Text, 9))
   temp_llave!PED_TIPMOV = 500
   temp_llave!PED_TRANSP = Val(Right(txtagencia.Text, 6))
   temp_llave("Ped_DESCTO") = dDouble(grid_fac.TextMatrix(fila, 20)) 'PORCENTAJE
   temp_llave("Ped_DESCTO_PRE") = dDouble(grid_fac.TextMatrix(fila, 19)) 'VALOR
   temp_llave.Update
pasa:
   fila = fila + 1
   WS_NRO_MOV = WS_NRO_MOV + 1
   If fila >= FORM_PED.grid_fac.Rows Then
      FLAG = True
   End If
  
Loop
FLAG_MODI = 0
Return

Screen.MousePointer = 1


Exit Sub
End If
Dim wser As String
Dim wnumfac As String

cmdIngreso.Caption = "&Grabar"
f1.Enabled = True
ESTADO.Enabled = True
LIMPIA_DATOS
FORM_PED.MODIF.Enabled = False
CABE_MAN
WMODO = "I"
PSTEMP_MAYOR(0) = LK_CODCIA
PSTEMP_MAYOR(1) = Val(tserie.Text)
temp_mayor.Requery
If temp_mayor.EOF Then
 wnumfac = 1
Else
 wnumfac = Val(Nulo_Valor0(temp_mayor!PED_NUMFAC)) + 1
End If
txtdoc.Text = wnumfac
'txtagencia.text = Nulo_Valors(par_llave!PAR_AGE_EMP)
If codpro.ListCount > 0 Then codpro.ListIndex = 0
fecha.Caption = Format(LK_FECHA_DIA, "dd/mm/yyyy")
moneda.ListIndex = 0
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
  FORM_PED.barra.Visible = False
  Screen.MousePointer = 0
  grid_fac.SetFocus
Else
  MsgBox Err.Description, 48, Pub_Titulo
End If

End Sub



Private Sub cmdmostrar_Click()
If Val(txtdias.Text) = 0 And Val(tpromedio.Text) = 0 Then
  GoTo pasa
End If


If Val(txtdias.Text) <= 0 Then
  MsgBox "Definir cantidad de dias para la orden de compra.", 48, Pub_Titulo
  txtdias.SetFocus
  Exit Sub
End If
If Val(tpromedio.Text) <= 0 Then
  MsgBox "Definir dias para el Promedio.", 48, Pub_Titulo
  tpromedio.SetFocus
  Exit Sub
End If
pasa:
fraListar.Visible = True
lisarti.Clear
opcomo(0).SetFocus
End Sub

Private Sub cmdpasar_Click()
Dim fx As Integer
Dim WCODIGO As Currency
Dim wcanti As Currency
Dim costo As Currency
Dim wdia, wpro As Integer
Dim wcalculo As Currency
Dim stockact As Currency
Dim stockmin As Currency
Dim tdias As Integer
Dim countartsel As Integer
Dim i As Integer

    If checia.Value = 1 Then
     LOC_CODCIA = "02"
    Else
     LOC_CODCIA = LK_CODCIA
    End If
    For i = 0 To lisarti.ListCount - 1
        If lisarti.Selected(i) = True Then
            countartsel = countartsel + 1
        End If
    Next i
    If countartsel = 0 Then
        MsgBox "Seleccione al menos un articulo", vbInformation, Pub_Titulo
        Exit Sub
    End If
    
    
    tdias = Val(txtdias.Text)
    wdia = Val(txtdias.Text)
    wpro = Val(tpromedio.Text)
If wpro = 0 And wdia = 0 Then GoTo pasa
If wdia <= 0 Then
 MsgBox "cantidad de Dias. No Procede..", 48, Pub_Titulo
 Exit Sub
End If
If wpro <= 0 Then
  MsgBox "Dias para el Promedio. No Procede..", 48, Pub_Titulo
  Exit Sub
End If
If wdia > wpro Then
  MsgBox "Cantidad de Dias no procede en el rango de promedio..", 48, Pub_Titulo
  Exit Sub
End If
pasa:
If lisarti.ListCount = 0 Then
 MsgBox "No Procede...", 48, Pub_Titulo
 Exit Sub
End If
PB.Visible = True
PB.Min = 0
PB.max = lisarti.ListCount
PB.Value = 0
SQ_OPER = 1
pu_codcia = LOC_CODCIA
wcalculo = 0
PUB_SECUEN = 0
For fila = 0 To lisarti.ListCount - 1
 If grid_fac.Rows >= Val(imax.Text) + 2 Then
     MsgBox "Llego al Maximo de Item.. Generar otra Orden de Compra", 48, Pub_Titulo
     GoTo TERMINAR
 End If
 PB.Value = PB.Value + 1
 lisarti.ListIndex = fila
 If lisarti.Selected(fila) = True Then
    PUB_KEY = Val(Right(lisarti, 8))
    PUB_CODART = Val(Right(lisarti, 8))
    SQ_OPER = 1
    LEER_ART_LLAVE
    stockact = 0
    If LISCIA.ListCount > 0 Then
     For fx = 0 To LISCIA.ListCount - 1
       LISCIA.ListIndex = fx
       If LISCIA.Selected(fx) Then
         pu_codcia = Left(LISCIA.Text, 2)
         LEER_ARM_LLAVE
         If Not arm_llave.EOF Then
            stockact = stockact + Nulo_Valor0(arm_llave!ARM_STOCK)
         Else
            stockact = 0
         End If
       End If
      Next fx
    Else
      pu_codcia = LK_CODCIA
      LEER_ARM_LLAVE
    End If
    pu_codcia = LOC_CODCIA
    SQ_OPER = 2
    LEER_PRE_LLAVE
    pre_mayor.MoveLast
    costo = 0
    WCODIGO = Val(Right(lisarti, 8))
    If wpro = 0 And wdia = 0 Then
      wcanti = 0
    Else
      wcanti = PROCESA_CANTIDAD(WCODIGO, tdias)
    End If
    If Not art_LLAVE.EOF Then
      costo = Nulo_Valor0(arm_llave!ARM_COSTO_ULT)
      'stockact = Nulo_Valor0(arm_llave!ARM_STOCK)
      stockmin = Nulo_Valor0(art_LLAVE!ART_STOCK_MIN)
      wcalculo = wcanti + stockmin - stockact
      If cheneg.Value = 1 And wcalculo < 0 Then GoTo sigueart
      If wpro = 0 And wdia = 0 Then
          wcalculo = 0
      End If
    End If
    grid_fac.Rows = grid_fac.Rows + 1
    grid_fac.RowHeight(grid_fac.Rows - 1) = 285
    grid_fac.TextMatrix(grid_fac.Rows - 1, 0) = "'" & Trim(art_LLAVE!art_nombre)  ' Trim(Mid(lisarti.Text, 11, 30))
    grid_fac.TextMatrix(grid_fac.Rows - 1, 1) = Trim(art_LLAVE!art_alterno)
    grid_fac.TextMatrix(grid_fac.Rows - 1, 2) = Format(wcalculo / pre_mayor!PRE_EQUIV, "#######")
    grid_fac.TextMatrix(grid_fac.Rows - 1, 3) = pre_mayor!pre_unidad
    If chesin.Value = 0 Then
      grid_fac.TextMatrix(grid_fac.Rows - 1, 4) = Format(costo, "0.00")
      grid_fac.TextMatrix(grid_fac.Rows - 1, 5) = Format((Val(grid_fac.TextMatrix(grid_fac.Rows - 1, 4)) * Val(grid_fac.TextMatrix(grid_fac.Rows - 1, 2))), "####0.00")
    End If
    grid_fac.TextMatrix(grid_fac.Rows - 1, 6) = Format(wcanti / pre_mayor!PRE_EQUIV, "#######")
    grid_fac.TextMatrix(grid_fac.Rows - 1, 7) = Format(stockact / pre_mayor!PRE_EQUIV, "#######")
    grid_fac.TextMatrix(grid_fac.Rows - 1, 8) = Format(stockmin / pre_mayor!PRE_EQUIV, "#######")
    grid_fac.TextMatrix(grid_fac.Rows - 1, 9) = PUB_KEY
    grid_fac.TextMatrix(grid_fac.Rows - 1, 10) = wcalculo
    grid_fac.TextMatrix(grid_fac.Rows - 1, 11) = costo
    grid_fac.TextMatrix(grid_fac.Rows - 1, 12) = pre_mayor!pre_secuencia

    grid_fac.TextMatrix(grid_fac.Rows - 1, 14) = wcanti
    grid_fac.TextMatrix(grid_fac.Rows - 1, 15) = stockact
    grid_fac.TextMatrix(grid_fac.Rows - 1, 16) = stockmin
    grid_fac.TextMatrix(grid_fac.Rows - 1, 17) = pre_mayor!PRE_EQUIV
    grid_fac.TextMatrix(grid_fac.Rows - 1, 18) = pre_mayor!pre_PESO
    grid_fac.ColWidth(12) = pre_mayor!pre_secuencia
    
 End If
sigueart:
Next fila

TERMINAR:
PB.Visible = False
suma_grid
lisarti.Clear
End Sub

Private Sub cmdproorden_Click()

Dim ws_conta As Integer
Dim WS_NOMCLI As String * 20
Dim WS_NOMBAN As String * 20
Dim WS_SALDO As Currency
Dim WS_MONEDA As String * 1
Dim ww_moneda
Dim ws_mensaje
Dim WS_FBG, WS_LARGO, ws_nomche, ws_codclie
Dim WS_SALDO_ING As Currency
Dim WS_SALDO_SAL As Currency
Dim todos_cli

Dim wsFECHA1
Dim ws_codcia As String
Dim xcuenta As Integer
Dim docus As String
Dim docus2 As String
Dim wmonto As Currency
Dim WINGRESOS As Currency
Dim TOT_INGRESOS As Currency
Dim WSALDO_CAJA As Currency



wsFECHA1 = LK_FECHA_DIA

DoEvents
DoEvents
GoSub WEXCEL

xl.Worksheets(1).Activate
WS_SALDO = 0
xcuenta = 0
f1 = 5


WSALDO_CAJA = 0
f1 = f1 + 1
ws_conta = 0

VAMOS:
WS_SALDO = 0
WSALDO_CAJA = 0
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
Else
  PUB_CODCIA = LK_CODCIA
End If


SQ_OPER = 1
LEER_PAR_LLAVE
WS_SALDO = par_llave!PAR_SALDO_CAJA_ayer
WSALDO_CAJA = WSALDO_CAJA + WS_SALDO
xl.Cells(f1, 2) = "Saldo Anterior:"
xl.Cells(f1, 5) = WS_SALDO
'all_llave.MoveNext
pub_cadena = "SELECT * FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ? AND ALL_FBG = ?  AND ALL_TIPMOV = 10 AND ALL_SIGNO_CAJA = 1  AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_DIA, ALL_FBG, ALL_NUMSER, ALL_NUMFAC"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT PAR_NOMBRE FROM PARGEN WHERE PAR_CODCIA = ? "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurValues)

xcuenta = 1
f1 = 6
f1 = f1 + 1
xl.Cells(f1, 1) = "MAS  : INGRESOS"
f1 = f1 + 1
xl.Cells(f1, 1) = 1
xl.Cells(f1, 2) = "VENTAS AL CONTADO:"
WINGRESOS = 0
TOT_INGRESOS = 0
For fila = 1 To 30
   ws_codcia = Mid(Trim(GEN!gen_ART_CIAS), xcuenta, 2)
   If Trim(ws_codcia) = "00" Then GoTo OTRA_CIA
   If Trim(ws_codcia) = "" Then Exit For
    WINGRESOS = 0
    PS_REP01(0) = ws_codcia
    PS_REP01(1) = wsFECHA1
    PS_REP02(0) = ws_codcia
    llave_rep02.Requery
    f1 = f1 + 1
    xl.Cells(f1, 2) = Trim(llave_rep02!PAR_NOMBRE)
    ' SOLO PARA FACTURAS
    PS_REP01(2) = "F"
    llave_rep01.Requery
    wmonto = 0
    If Not llave_rep01.EOF Then
       docus = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
      Do Until llave_rep01.EOF
        wmonto = wmonto + llave_rep01!ALL_IMPORTE_AMORT
        docus2 = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
        llave_rep01.MoveNext
      Loop
      f1 = f1 + 1
      ' xl.Cells(f1, 3) = "F/. " & docus & " AL " & docus2
      xl.Cells(f1, 3) = "FACTURAS"
      xl.Cells(f1, 4) = wmonto
      WINGRESOS = WINGRESOS + wmonto
    End If
    'SOLO PARA BOLETAS
    PS_REP01(2) = "B"
    llave_rep01.Requery
    wmonto = 0
    If Not llave_rep01.EOF Then
       docus = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
      Do Until llave_rep01.EOF
        wmonto = wmonto + llave_rep01!ALL_IMPORTE_AMORT
        docus2 = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
        llave_rep01.MoveNext
      Loop
      f1 = f1 + 1
      ' xl.Cells(f1, 3) = "B/. " & docus & " AL " & docus2
      xl.Cells(f1, 3) = "BOLETAS"
      xl.Cells(f1, 4) = wmonto
      WINGRESOS = WINGRESOS + wmonto
    End If
    ' SOLO PARA GUIAS
    PS_REP01(2) = "G"
    llave_rep01.Requery
    wmonto = 0
    If Not llave_rep01.EOF Then
       docus = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
      Do Until llave_rep01.EOF
        wmonto = wmonto + llave_rep01!ALL_IMPORTE_AMORT
        docus2 = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
        llave_rep01.MoveNext
      Loop
      f1 = f1 + 1
      'xl.Cells(f1, 3) = "G/. " & docus & " AL " & docus2
      xl.Cells(f1, 3) = "GUIAS"
      xl.Cells(f1, 4) = wmonto
      WINGRESOS = WINGRESOS + wmonto
    End If
    ' SOLO PARA VENTAS ADMINISTRADORES P
    PS_REP01(2) = "P"
    llave_rep01.Requery
    wmonto = 0
    If Not llave_rep01.EOF Then
       docus = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
      Do Until llave_rep01.EOF
        wmonto = wmonto + llave_rep01!ALL_IMPORTE_AMORT
        docus2 = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
        llave_rep01.MoveNext
      Loop
      f1 = f1 + 1
      xl.Cells(f1, 3) = "P/. " & docus & " AL " & docus2
      xl.Cells(f1, 3) = "ADMINISTRACION"
      xl.Cells(f1, 4) = wmonto
      WINGRESOS = WINGRESOS + wmonto
    End If
      f1 = f1 + 1
      xl.Cells(f1, 4) = WINGRESOS
      TOT_INGRESOS = TOT_INGRESOS + WINGRESOS
OTRA_CIA:
   xcuenta = xcuenta + 2
 Next fila
 f1 = f1 + 1
 xl.Cells(f1, 5) = TOT_INGRESOS
 WSALDO_CAJA = WSALDO_CAJA + TOT_INGRESOS

DoEvents

DoEvents
xcuenta = 1
xl.Cells(2, 2) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
xl.Cells(2, 5) = "'" & Format(wsFECHA1, "dd mmm yyyy")
xl.DisplayAlerts = False
xl.Worksheets("Hoja1").Range("A1:X51").Locked = True
xl.Worksheets("Hoja1").Protect PUB_CLAVE
xl.Application.Visible = True
DoEvents
Set xl = Nothing
Screen.MousePointer = 0


Exit Sub


WEXCEL:
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  
  DoEvents
  
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\OFFICE\CAJA.xls", 0, True, 4, PUB_CLAVE, PUB_CLAVE
Return

Exit Sub

'LLENA_VALOR:
'For I = QJ To WDIF
'  If I >= 3 Then
'    xl.Cells(F1 + 7, I) = Format(LOC_VALOR, "0.00")
'  End If
'Next I
'Return

Exit Sub
CANCELA:
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FORM_PED
Exit Sub



End Sub

Private Sub cmdSelecAll_Click(Index As Integer)
Dim i As Integer
    For i = 0 To lisarti.ListCount - 1
        If Index = 0 Then
            lisarti.Selected(i) = True
        ElseIf Index = 1 Then
            lisarti.Selected(i) = False
        End If
    Next i
End Sub

Private Sub codpro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Azul txtdias, txtdias
End If
End Sub

Private Sub codpro_LostFocus()

SQ_OPER = 1
pu_codclie = Val(Right(codpro.Text, 8))
pu_cp = "P"
pu_codcia = LK_CODCIA
LEER_CLI_LLAVE
'' fpago.Text = cli_llave!CLI_AUTOAVALUO
If cli_llave.EOF Then Exit Sub
txtContacto.Text = Nulo_Valors(cli_llave!CLI_NOMBRE_ESPOSA)


End Sub


Private Sub fami_Click()
Dim wpos As Integer
Dim WFAMI2 As Integer
If opcomo(1) = True Then
    lisarti.Clear
    If Trim(fami.Text) = "" Then
     subfami.Clear
     Exit Sub
    End If
    Screen.MousePointer = 11
    wpos = fami.ListIndex
    WFAMI2 = Val(Trim(Right(fami.Text, 6)))
    LLENADO_SUBFAM WFAMI2
    On Error GoTo sigue
    subfami.ListIndex = wpos
    Screen.MousePointer = 0
    Exit Sub
sigue:
    Resume Next
End If
If opcomo(0) = True Then
 lisarti.Clear
 LLENA_ARTI Val(Right(fami.Text, 6)), 0
End If

End Sub

Private Sub Form_Load()
Dim xcuenta As Integer
'On Error GoTo SALE
checia.Visible = False
'If LK_EMP = "3AA" Then
' checia.Visible = True
'End If

Wsec = 0
LOC_CANCELA = 0
fila = 0
wfila_act = 0
WSELE = ""
Dim ws_indice As Integer
Dim cade
WMODO = ""

pub_cadena = "SELECT * FROM CLIENTES WHERE CLI_CP = 'P'  AND CLI_CODCIA = ? ORDER BY CLI_NOMBRE"
Set PSPRO_V = CN.CreateQuery("", pub_cadena)
PSPRO_V.rdoParameters(0) = " "
Set PRO_V = PSPRO_V.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PSPRO_V(0) = LK_CODCIA
PRO_V.Requery
codpro.Clear
Do Until PRO_V.EOF
    codpro.AddItem PRO_V!CLI_NOMBRE & String(60, " ") & PRO_V!cli_codclie
    PRO_V.MoveNext
Loop

pub_cadena = "SELECT * FROM TRANSPORTE ORDER  BY TRN_NOMBRE"
Set PSPRO_V = CN.CreateQuery("", pub_cadena)
Set PRO_V = PSPRO_V.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PRO_V.Requery
txtagencia.Clear
Do Until PRO_V.EOF
    txtagencia.AddItem PRO_V!TRN_NOMBRE & String(80, " ") & PRO_V!TRN_KEY
    PRO_V.MoveNext
Loop


pub_cadena = "SELECT SUM(FAR_CANTIDAD)AS CANTIDAD FROM FACART WHERE (FAR_CODCIA = ? or FAR_CODCIA = ? or FAR_CODCIA = ?) AND FAR_CODART = ? and (FAR_FECHA >= ? AND FAR_FECHA <= ?) AND FAR_TIPMOV = 10  AND FAR_ESTADO  <> 'E' GROUP BY FAR_CODART"
Set PSLOC_WARTI = CN.CreateQuery("", pub_cadena)
PSLOC_WARTI.rdoParameters(0) = 0
PSLOC_WARTI.rdoParameters(1) = 0
PSLOC_WARTI.rdoParameters(2) = 0
PSLOC_WARTI.rdoParameters(3) = 0
PSLOC_WARTI.rdoParameters(4) = LK_FECHA_DIA
PSLOC_WARTI.rdoParameters(5) = LK_FECHA_DIA
Set llave_sum_arti = PSLOC_WARTI.OpenResultset(rdOpenKeyset, rdConcurValues)



pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_NUMSER = ? AND PED_TIPMOV= 500  ORDER BY  PED_NUMFAC DESC "
Set PSTEMP_MAYOR = CN.CreateQuery("", pub_cadena)
PSTEMP_MAYOR.rdoParameters(0) = " "
PSTEMP_MAYOR.rdoParameters(1) = 0

PSTEMP_MAYOR.MaxRows = 1
Set temp_mayor = PSTEMP_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_CODCIA= ? AND PED_NUMSER = ? AND PED_NUMFAC= ? AND PED_TIPMOV = 500 ORDER BY PED_CODCIA"
Set PSTEMP_LLAVE = CN.CreateQuery("", pub_cadena)
PSTEMP_LLAVE.rdoParameters(0) = " "
PSTEMP_LLAVE.rdoParameters(1) = 0
PSTEMP_LLAVE.rdoParameters(2) = 0
Set temp_llave = PSTEMP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
temp_llave.Requery


pub_cadena = "SELECT ART_KEY  FROM ARTI WHERE ART_CODCIA = ? AND ART_ALTERNO =?"
Set PS_REP04 = CN.CreateQuery("", pub_cadena)
PS_REP04.rdoParameters(0) = ""
PS_REP04.rdoParameters(1) = ""
Set llave_rep04 = PS_REP04.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER >=? AND COV_FECHA_VOUCHER <=?  ORDER BY COV_CODCIA, COV_NRO_VOUCHER, COV_DH" ', COV_NRO_MOV"
Set PSCOV_MAYOR = CN.CreateQuery("", pub_cadena)
PSCOV_MAYOR.rdoParameters(0) = " "
PSCOV_MAYOR.rdoParameters(1) = LK_FECHA_DIA
PSCOV_MAYOR.rdoParameters(2) = LK_FECHA_DIA
Set cov_mayor = PSCOV_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ?  AND COV_FECHA_VOUCHER=? AND COV_NRO_VOUCHER = ? AND COV_NRO_MOV = ? ORDER BY COV_CODCIA, COV_FECHA_VOUCHER, COV_NRO_VOUCHER, COV_NRO_MOV"
Set PSCOV_LLAVE = CN.CreateQuery("", pub_cadena)
PSCOV_LLAVE(0) = " "
PSCOV_LLAVE(1) = LK_FECHA_DIA
PSCOV_LLAVE(2) = 0
PSCOV_LLAVE(3) = 0
Set cov_llave = PSCOV_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM COPARAM WHERE COP_CODCIA = ?"
Set PSCOP_LLAVE = CN.CreateQuery("", pub_cadena)
PSCOP_LLAVE.rdoParameters(0) = " "
Set cop_llave = PSCOP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
If LK_EMP = "3AA" Then
   Combo1.AddItem "AV. PROLONGACION UNION N. 2218 - TRUJILLO"
   Combo1.AddItem "AV. SANCHEZ CERRO N. 1675 - PIURA"
   Combo1.ListIndex = 0
End If
''If LK_EMP_PTO = "A" Then
'' PSCOP_LLAVE.rdoParameters(0) = "00"
''Else
 PSCOP_LLAVE.rdoParameters(0) = par_llave!PAR_CIACON
'End If

cop_llave.Requery
PUB_CODCIA = LK_CODCIA
LLENA_COMBO fpago, 11
fila = 0
DoEvents
LIMPIA_DATOS
CABE_MAN
tserie.Text = 1


 LISCIA.Visible = True
 LISCIA.Clear
 xcuenta = 0
 For fila = 1 To 30 Step 2
   PUB_CODCIA = Mid(Trim(par_llave!par_art_cias), fila, 2)
   If Trim(PUB_CODCIA) = "" Then Exit For
   xcuenta = xcuenta + 1
   PSPAR_MULTI(0) = PUB_CODCIA
   par_multi.Requery
   LISCIA.AddItem PUB_CODCIA & " - " & Trim(par_multi!PAR_NOMBRE)
 Next fila
 
 For fila = 0 To LISCIA.ListCount - 1
  LISCIA.ListIndex = fila
  If Left(LISCIA.Text, 2) = LK_CODCIA Then LISCIA.Selected(fila) = True
 Next fila





Exit Sub
SALE:
MsgBox "Depurar: " & Err.Description, 48, Pub_Titulo
Resume Next
End Sub

Private Sub fpago_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 45 Then
 Exit Sub
End If
Dim wpos
wpos = fpago.ListIndex
PUB_TIPREG = Mid(fpago.ToolTipText, 13, Len(fpago.ToolTipText))
PUB_CODCIA = LK_CODCIA
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If
Load FrmDatArti
FrmDatArti.Caption = "GRUPOS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENA_COMBO fpago, 11
On Error GoTo sigue
fpago.ListIndex = wpos
On Error GoTo 0
fpago.SetFocus
SendKeys "%{up}"
Exit Sub
sigue:
Resume Next


End Sub

Private Sub grid_fac_DblClick()
    If Val(grid_fac.TextMatrix(grid_fac.Row, 9)) <> 0 Then
        PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 9))
        frmchart.ArtiFacart
    End If
End Sub

Private Sub grid_fac_GotFocus()
    fraListar.Visible = False
End Sub

Private Sub grid_fac_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then
        If Trim(grid_fac.TextMatrix(grid_fac.Row, 21)) = "" Then
            grid_fac.TextMatrix(grid_fac.Row, 21) = "X"
            BackColorRow grid_fac.Row, grid_fac, &HC0C0FF
        Else
            grid_fac.TextMatrix(grid_fac.Row, 21) = " "
            BackColorRow grid_fac.Row, grid_fac, &HFFFFFF
        End If
    End If
End Sub

Private Sub grid_fac_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, WC
Static CONS
If KeyAscii <> 13 Then Exit Sub
If grid_fac.Rows <= 2 Then Exit Sub
If grid_fac.COL = 1 Then Exit Sub
If grid_fac.COL >= 6 And grid_fac.COL <> 20 Then Exit Sub


'GoTo leer
'Exit Sub
'End If
If WMODO = "I" Then
    'If Trim(grid_fac.TextMatrix(grid_fac.Row, 0)) = "" Then Exit Sub
    'If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) <> "" And grid_fac.Col = 2 Or grid_fac.Col = 3 Then GoTo leer
    'If Trim(grid_fac.TextMatrix(grid_fac.Row, 8)) <> "0" Then Exit Sub
End If


'If grid_fac.Col = 1 And WMODO = "I" Then
'   a = Val(grid_fac.TextMatrix(grid_fac.Row - 1, 0))
'   a = a + 1
'  grid_fac.TextMatrix(grid_fac.Row, 0) = a
'End If

If WMODO = "I" Or FLAG_MODI = 1 Then
  If grid_fac.COL = 3 Then
    UNIDAD.Left = grid_fac.Left + grid_fac.CellLeft
    UNIDAD.Width = grid_fac.CellWidth
    'UNIDAD.Height = grid_fac.CellHeight
    UNIDAD.Top = ESTADO.Top + grid_fac.Top + grid_fac.CellTop - 1200 '480
    SQ_OPER = 2
    pu_codcia = LK_CODCIA
    PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 9))
    LEER_PRE_LLAVE
    UNIDAD.Clear
    UNIDAD.Visible = True
    Do Until pre_mayor.EOF
     UNIDAD.AddItem Trim(pre_mayor!pre_unidad) & String(30, " ") & pre_mayor!pre_secuencia
     pre_mayor.MoveNext
    Loop
    'wfila_act = grid_fac.Row
    On Error GoTo pasa
    UNIDAD.ListIndex = Val(grid_fac.TextMatrix(grid_fac.Row, 12))
    On Error GoTo 0
    UNIDAD.Visible = True
    
    UNIDAD.SetFocus
    SendKeys "%{up}"
     Exit Sub
  End If
    TEXTOVAR.Left = grid_fac.Left + grid_fac.CellLeft
    TEXTOVAR.Width = grid_fac.CellWidth
    TEXTOVAR.Height = grid_fac.CellHeight
    TEXTOVAR.Top = grid_fac.Top + grid_fac.CellTop - 10 '- 2025 '480ESTADO.Top +
    TEXTOVAR.Text = grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL)
    wfila_act = grid_fac.Row
    TEXTOVAR.Visible = True
    Azul TEXTOVAR, TEXTOVAR
    TEXTOVAR.SetFocus
End If
Exit Sub
pasa:
Resume Next
End Sub

Private Sub grid_fac_KeyUp(KeyCode As Integer, Shift As Integer)
Dim WC
Dim a, WF As Integer
Dim tf, t, tC
Dim SALE As Boolean
Dim i As Integer
Dim NroRows As Integer
If KeyCode = 46 Then
If grid_fac.Rows <= 2 Then Exit Sub
If grid_fac.Rows <= 3 Then
    pub_mensaje = MsgBox("Quitar el Producto para la Orden de Compra ", vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
    If pub_mensaje = vbNo Then
      grid_fac.SetFocus
      Exit Sub
    End If
    CABE_MAN
Else
   pub_mensaje = MsgBox("Quitar el Producto para la Orden de Compra ", vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
   If pub_mensaje = vbNo Then
     grid_fac.SetFocus
     Exit Sub
   Else
   '  grid_fac.RowHeight(grid_fac.Row) = 1
   NroRows = grid_fac.Rows - 2
   For i = NroRows To 1 Step -1
        If grid_fac.TextMatrix(i + 1, 21) = "X" Then
            grid_fac.RemoveItem (i + 1)
            'NroRows = NroRows - 1
        End If
   Next i
   
   grid_fac.Row = grid_fac.Row
   grid_fac.Refresh
   suma_grid
   grid_fac.SetFocus
   End If
End If
End If
'grid_fac.SetFocus
Exit Sub



End Sub


Private Sub i_codart2_Change()
If i_codart2.Text = "" Then
  VAR_ACTIVAR = 0
End If

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
'  LISTVIEW1.Visible = False
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
    If art_LLAVE!art_flag_stock <> "M" Then
       MsgBox "Producto no es Mercaderia.", 48, Pub_Titulo
       Azul i_codart2, i_codart2
       GoTo fin
    End If
    WCOD_ORIGINAL = art_LLAVE!ART_KEY
    'i_codart2.text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
    lisarti.AddItem art_LLAVE!art_alterno & " " & Trim(art_LLAVE!art_nombre) & String(60, " ") & art_LLAVE!ART_KEY
    lisarti.Selected(lisarti.ListCount - 1) = True
    i_codart2.Text = ""
    ListView1.Visible = False
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
     If art_llave_alt!art_flag_stock <> "M" Then
       MsgBox "Producto no es Mercaderia.", 48, Pub_Titulo
       Azul i_codart2, i_codart2
       GoTo fin
     End If
     WCOD_ORIGINAL = art_llave_alt!ART_KEY
     'i_codart2.text = Trim(art_llave_alt!ART_NOMBRE)
     lisarti.AddItem art_llave_alt!art_alterno & " " & Trim(art_llave_alt!art_nombre) & String(60, " ") & art_llave_alt!ART_KEY
     lisarti.Selected(lisarti.ListCount - 1) = True
     ListView1.Visible = False
     i_codart2.Text = ""
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
      If art_LLAVE!art_flag_stock <> "M" Then
       MsgBox "Producto no es Mercaderia.", 48, Pub_Titulo
       Azul i_codart2, i_codart2
       GoTo fin
      End If
      WCOD_ORIGINAL = art_LLAVE!ART_KEY
      lisarti.AddItem art_LLAVE!art_alterno & " " & Trim(art_LLAVE!art_nombre) & String(60, " ") & art_LLAVE!ART_KEY
      lisarti.Selected(lisarti.ListCount - 1) = True
      i_codart2.Text = ""
      ListView1.Visible = False
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
      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_CODCIA = '" & LK_CODCIA & "' AND ART_CALIDAD = 1 AND ART_FLAG_STOCK = 'M' AND ART_ALTERNO BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
    Else
      numarchi = 0
      'archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_CODCIA = '" & LK_CODCIA & "' AND ART_CALIDAD = 1 AND ART_FLAG_STOCK = 'M' AND ART_NOMBRE BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
      archi = "SELECT ARTI.ART_KEY,ARTI.ART_CODCIA, ARTI.ART_NOMBRE, ARTI.ART_ALTERNO, ARTICULO.ARM_STOCK, PRECIOS.PRE_EQUIV, TABLAS.TAB_NOMLARGO AS DIVI, TABLAS_1.TAB_NOMLARGO AS LINEA "
        archi = archi & "FROM ARTI INNER JOIN ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA INNER JOIN PRECIOS ON ARTI.ART_KEY = PRECIOS.PRE_CODART AND ARTI.ART_CODCIA = PRECIOS.PRE_CODCIA INNER JOIN TABLAS ON ARTI.ART_CODCIA = TABLAS.TAB_CODCIA AND ARTI.ART_FAMILIA = TABLAS.TAB_NUMTAB INNER JOIN TABLAS TABLAS_1 ON ARTI.ART_CODCIA = TABLAS_1.TAB_CODCIA AND ARTI.ART_SUBFAM = TABLAS_1.TAB_NUMTAB "
        archi = archi & "WHERE (TABLAS.TAB_TIPREG = 122) AND (TABLAS_1.TAB_TIPREG = 123) AND (PRECIOS.PRE_FLAG_UNIDAD = 'A') AND ARTI.ART_CODCIA = '" & LK_CODCIA & "' AND ARTI.ART_NOMBRE BETWEEN '" & Trim(i_codart2.Text) & "%' AND  '" & var & "' ORDER BY ARTI.ART_NOMBRE"
    End If
   ' If Len(I_CODART2.text) > 1 And ListView1.ListItems.count = 0 Then
   ' Else
     PROC_LISVIEW ListView1
   ' End If
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
' ListView1.Visible = False
End Sub



Private Sub Imprimir_Click()
On Error GoTo FINTODO
Dim wser As String * 3
Dim WSRUTA As String
Dim wRuta As String
Dim rmoneda As String * 1
wRuta = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\"
Reportes.Connect = PUB_ODBC
Reportes.Destination = crptToWindow  '= crptToPrinter
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 800
Reportes.WindowHeight = 490
Reportes.Formulas(1) = ""
PUB_NETO = Val(txttotal.Text)
PUB_FECHA = fecha.Caption
PU_NUMSER = Val((tserie.Text))

If Left(moneda.Text, 1) = "D" Then
   rmoneda = "D"
Else
   rmoneda = "S"
End If

PU_NUMFAC = Val((txtdoc.Text))
Reportes.Formulas(1) = "SON_EFECTIVO=  'SON: " & CONVER_LETRAS(PUB_NETO, rmoneda) & "'"
Reportes.WindowTitle = "ORDEN DE COMPRA  :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
Reportes.ReportFileName = wRuta + "ORDEN.RPT"
wser = PU_NUMSER
pub_cadena = "{PEDIDOS.PED_ESTADO} = 'N' AND {PEDIDOS.PED_CODCIA} = '" & LK_CODCIA & "' AND {PEDIDOS.PED_NUMSER}= '" & wser & "' AND {PEDIDOS.PED_NUMFAC} = " & PU_NUMFAC
Reportes.SelectionFormula = pub_cadena
Reportes.WindowTitle = Reportes.WindowTitle & " Archivo: " & Trim(Reportes.ReportFileName)
On Error GoTo accion
Reportes.Action = 1
On Error GoTo 0
Exit Sub

FINTODO:
accion:
 MsgBox Err.Description, 48, Pub_Titulo
 MsgBox "Reintente Nuevamente ..", 48, Pub_Titulo

End Sub

Private Sub lisarti_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer
If KeyCode = 113 Then
  For i = 0 To lisarti.ListCount - 1
   lisarti.ListIndex = i
   lisarti.Selected(i) = True
  Next i
End If
If KeyCode = 114 Then
  For i = 0 To lisarti.ListCount - 1
   lisarti.ListIndex = i
   lisarti.Selected(i) = False
  Next i
End If
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
If KeyAscii = 27 Then
 ListView1.Visible = False
 i_codart2.Text = ""
 i_codart2.SetFocus
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

Private Sub modif_Click()
Dim ws_tot_debe, ws_tot_haber As Currency
Dim er As rdoError
Dim pub_mensaje As String
Const ingre = 2
Const MODIF = 1
Dim N As Integer
Dim LOC_SALDO_CAR As Currency
Dim FLAG As Boolean
Dim pub_mensaje_err As String
Dim WS_NRO_MOV, ws_nro_voucher As Long
Dim w_dh  As String
FLAG_MODI = 1
tserie.Enabled = True
txtdoc.Enabled = True
FORM_PED.MODIF.Enabled = False

f1.Enabled = True
cmdIngreso.Caption = "&Grabar"
ESTADO.Enabled = True
tserie.SetFocus


End Sub

Private Sub moneda_Click()
If Left(moneda.Text, 1) = "S" Then
 lblMoneda.Caption = "S/."
Else
 lblMoneda.Caption = "US$."
End If
End Sub

Private Sub MSChart1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 FORM_PED.MSChart1.Visible = False
End If
End Sub

Private Sub MSChart1_LostFocus()
 FORM_PED.MSChart1.Visible = False
End Sub

Private Sub opcomo_Click(Index As Integer)
Select Case Index
Case 0
  lblarti.Visible = False
  i_codart2.Visible = False
  lsubfami.Visible = False
  subfami.Visible = False
  subfami.Clear
  subfami.Enabled = False
  lfami.Visible = True
  fami.Clear
  fami.Visible = True
  fami.Enabled = True
  fami.Width = 7425
  
  PUB_CODCIA = LK_CODCIA
  larti.Top = 3200
  lisarti.Top = 3400
  lisarti.Height = 2700
  LLENADOS fami, 122
Case 1
  lblarti.Visible = False
  i_codart2.Visible = False
  lsubfami.Visible = True
  subfami.Visible = True
  subfami.Clear
  subfami.Enabled = True
  
  fami.Visible = True
  lfami.Visible = True
  fami.Clear
  fami.Enabled = True
  fami.Width = 3375
  
  larti.Top = 3200
  lisarti.Top = 3400
  lisarti.Height = 2700
  
  PUB_CODCIA = LK_CODCIA
  LLENADOS fami, 122
Case 2
  lblarti.Visible = False
  i_codart2.Visible = False
  fami.Visible = False
  lfami.Visible = False
  fami.Clear
  fami.Enabled = False
  subfami.Visible = False
  lsubfami.Visible = False
  subfami.Clear
  subfami.Enabled = False
  larti.Top = 500
  lisarti.Top = 700
  lisarti.Height = 5300
  LLENA_ARTI_PRO Val(Right(codpro.Text, 10))
Case 3
  fami.Visible = False
  lfami.Visible = False
  subfami.Visible = False
  lsubfami.Visible = False
  subfami.Clear
  subfami.Enabled = False
  
  lblarti.Visible = True
  lisarti.Clear
  larti.Top = 500
  lisarti.Top = 1000
  lisarti.Height = 5000
  i_codart2.Visible = True
  i_codart2.Text = ""
  i_codart2.SetFocus
End Select


End Sub

Private Sub opcomo_DblClick(Index As Integer)
opcomo_Click Index
End Sub

Private Sub salir_Click()
Unload FORM_PED
End Sub


Public Sub LIMPIA_DATOS()
codpro.ListIndex = -1
'tserie.text = ""
txtdoc.Text = ""
grid_fac.Clear
'txtagencia.text = ""
txtigv.Text = ""
txtvalorv.Text = ""
txttotal.Text = ""
lblpeso.Caption = ""

End Sub

Public Sub CABE_MAN()
grid_fac.Cols = 22
grid_fac.Rows = 2
grid_fac.Clear
fila = 0
grid_fac.ColWidth(0) = 3500
grid_fac.ColWidth(1) = 900
grid_fac.ColWidth(2) = 900
grid_fac.ColWidth(3) = 550
grid_fac.ColWidth(4) = 900
grid_fac.ColWidth(5) = 1000
grid_fac.ColWidth(6) = 800
grid_fac.ColWidth(7) = 800
grid_fac.ColWidth(8) = 900
grid_fac.ColWidth(9) = 0
grid_fac.ColWidth(10) = 0
grid_fac.ColWidth(11) = 0
grid_fac.ColWidth(12) = 0
grid_fac.ColWidth(13) = 0
grid_fac.ColWidth(14) = 0
grid_fac.ColWidth(15) = 0
grid_fac.ColWidth(16) = 0
grid_fac.ColWidth(17) = 0
grid_fac.ColWidth(18) = 0
'agregado mic
grid_fac.ColWidth(19) = 1200
grid_fac.ColWidth(20) = 1200
grid_fac.ColWidth(21) = 0

grid_fac.TextMatrix(0, 0) = "Articulo"
grid_fac.TextMatrix(0, 1) = "Codigo"
grid_fac.TextMatrix(0, 2) = "Cantidad"
grid_fac.TextMatrix(0, 3) = "Unidad"
grid_fac.TextMatrix(0, 4) = "Precios"
grid_fac.TextMatrix(0, 5) = "Sub Total"
grid_fac.TextMatrix(0, 6) = "Prom. Ventas"
grid_fac.TextMatrix(0, 7) = "Stock Actual"
grid_fac.TextMatrix(0, 8) = "Stock Min"
'agregado mic
grid_fac.TextMatrix(0, 19) = "Dscto"
grid_fac.TextMatrix(0, 20) = "Dscto %"

End Sub
Public Sub suma_grid()
'On Error GoTo SALE
Dim WF As Integer
Dim ww_subt As Currency
Dim SUM_PESO As Currency
WF = 2
Dim fx As Integer
fx = 1
SUM_H = 0
SUM_D = 0
SUM_PESO = 0
PUB_DESCTO = 0
Do While fx = 1
    If grid_fac.TextMatrix(WF, 0) <> "" Then
      If Val(grid_fac.TextMatrix(WF, 19)) <> 5 Then
         ww_subt = Format(Val(Val(grid_fac.TextMatrix(WF, 2)) * Val(grid_fac.TextMatrix(WF, 4))), "0.00")
         grid_fac.TextMatrix(WF, 5) = ww_subt
      Else
         ww_subt = Val(grid_fac.TextMatrix(WF, 5))
         If Val(grid_fac.TextMatrix(WF, 2)) <> 0 Then
         grid_fac.TextMatrix(WF, 4) = Format(Val(Val(grid_fac.TextMatrix(WF, 5)) / Val(grid_fac.TextMatrix(WF, 2))), "0.0000")
         End If
      End If
      SUM_H = SUM_H + ww_subt
      SUM_D = SUM_D + Val(grid_fac.TextMatrix(WF, 4))
      SUM_PESO = SUM_PESO + (Val(grid_fac.TextMatrix(WF, 18)) * Val(grid_fac.TextMatrix(WF, 2)))
      'agregado por mic
      'porcentaje
      grid_fac.TextMatrix(WF, 19) = ww_subt * Val(grid_fac.TextMatrix(WF, 20)) / 100
        PUB_DESCTO = PUB_DESCTO + Val(grid_fac.TextMatrix(WF, 19))
    
    End If
    WF = WF + 1
    If WF = grid_fac.Rows Then
        fx = 0
    Else
        If Trim(grid_fac.TextMatrix(WF, 0)) = "" Then fx = 0
    End If
Loop
   fila = WF - 1
   grid_fac.TextMatrix(1, 5) = Format(SUM_H, "####0.00")
   lblpeso.Caption = Format(SUM_PESO, "#,##0.00")
  ' txttotal.text = Format(SUM_H, "###,##0.00")
  ' txtigv.text = Format((SUM_H / LK_IGV), "###,##0.00")
   'txtvalorv.text = Format(Val(txttotal.text) - Val(txtigv.text), "###,##0.00")
   If LK_EMP = "HER" Then
    txtvalorv.Text = SUM_H
    txtigv.Text = Format((SUM_H - PUB_DESCTO) * ((LK_IGV / 100)), "0.00")
    txttotal.Text = Format((SUM_H - PUB_DESCTO) * (1 + (LK_IGV / 100)), "0.00") ' Val(Format(SUM_H, "#####0.00")) + Val(Format(txtigv.Text, "#####0.00"))
    txtdescto.Text = Format(PUB_DESCTO, "0.00")
   Else
    txtigv.Text = SUM_H - Format((SUM_H / (1 + (LK_IGV / 100))), "####0.00")
    txttotal.Text = Format(SUM_H, "#####0.00")
    txtvalorv.Text = Val(txttotal.Text) - Val(txtigv.Text)
   End If
   
   
Exit Sub
SALE:
cancelar_Click
'MsgBox "Verficar Importe.", 48, Pub_Titulo
'Resume Next
'If TEXTOVAR.Visible Then Azul3 TEXTOVAR, TEXTOVAR
End Sub
Public Sub suma_subtotal()
If WMODO = "I" Then Exit Sub

Dim WF As Integer
Dim WFIN As Integer
Dim WINI As Integer

Dim fx As Integer
Exit Sub
End Sub

Private Sub Consistencias(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyAscii As Integer)
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
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And car <> "." Then
          wsKeyAscii = 0
          Beep
          Exit Sub
        End If
    End If

End Sub

Public Sub CABE_ING()
grid_fac.Cols = 6
grid_fac.Rows = 3
grid_fac.Clear
grid_fac.MergeCells = 4
grid_fac.MergeCol(0) = True
grid_fac.MergeCol(1) = True
grid_fac.MergeCol(2) = True
grid_fac.MergeCol(3) = True
grid_fac.MergeCol(4) = False
grid_fac.MergeCol(5) = False
grid_fac.MergeRow(2) = False
grid_fac.RowHeight(0) = 285
grid_fac.RowHeight(1) = 285
grid_fac.RowHeight(2) = 285

fila = 0
grid_fac.ColWidth(0) = 400
grid_fac.ColWidth(1) = 1400
grid_fac.ColWidth(2) = 2500
grid_fac.ColWidth(3) = 0
grid_fac.ColWidth(4) = 1500
grid_fac.ColWidth(5) = 1500

grid_fac.TextMatrix(0, 0) = "Item"
grid_fac.TextMatrix(0, 1) = "Cuenta"
grid_fac.TextMatrix(0, 2) = "Descripcion"
grid_fac.TextMatrix(0, 3) = "Glosa"
grid_fac.TextMatrix(0, 4) = "Debe"
grid_fac.TextMatrix(0, 5) = "Haber"
grid_fac.TextMatrix(1, 0) = "Item"
grid_fac.TextMatrix(1, 1) = "Cuenta"
grid_fac.TextMatrix(1, 2) = "Descripcion"
grid_fac.TextMatrix(1, 3) = "Glosa"

'grid_fac.MergeCol
'grid_fac.MergeRow(2) = True



End Sub

Private Sub SALIR_GotFocus()
fraListar.Visible = False
End Sub

Private Sub subfami_Click()
If Trim(subfami.Text) = "" Then
Exit Sub
End If
LLENA_ARTI 0, Val(Right(subfami.Text, 6))
End Sub
Private Sub textovar_Change()
If grid_fac.COL = 1 Then
Else
 If grid_fac.COL = 2 Or grid_fac.COL = 20 Then
  grid_fac.Text = TEXTOVAR.Text
 Else
  grid_fac.Text = Format(TEXTOVAR.Text, "0.00")
 End If
 If grid_fac.COL = 5 Then grid_fac.TextMatrix(grid_fac.Row, 19) = grid_fac.COL
End If
suma_grid
suma_subtotal

End Sub

Private Sub TEXTOVAR_GotFocus()
temporal = grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL)
End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
'SOLO_DECIMAL TEXTOVAR, KeyAscii
If KeyAscii = 27 Then
  TEXTOVAR.Text = temporal
  TEXTOVAR.Visible = False
  grid_fac.SetFocus
  Exit Sub
End If
If grid_fac.COL = 2 Or grid_fac.COL = 4 Then Consistencias grid_fac, TEXTOVAR, KeyAscii
If KeyAscii <> 13 Then Exit Sub

TEXTOVAR.Visible = False
grid_fac.SetFocus

fin:

End Sub

Private Sub textovar_LostFocus()
'TEXTOVAR.Visible = False
If TEXTOVAR.Visible Then
   TEXTOVAR.Visible = False
   grid_fac.Row = wfila_act
'   grid_fac.SetFocus
   Exit Sub
End If

End Sub

Public Sub LLENADO_SUBFAM(wfami As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = 123
    PUB_CODCIA = LK_CODCIA
    PUB_CODART = wfami
    SQ_OPER = 3
    LEER_TAB_LLAVE
    subfami.ToolTipText = "TAB_TIPREG = 123"
    subfami.Clear
    Do While Not tab_menor.EOF
        subfami.AddItem tab_menor!tab_NOMLARGO & String(100, " ") & Trim(CStr(tab_menor!TAB_NUMTAB))
        CONTA = CONTA + 1
        tab_menor.MoveNext
    Loop
End Sub
Public Sub LLENADOS(cont As ListBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    cont.AddItem " "
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!tab_NOMLARGO & String(150, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub
Public Sub LLENA_COMBO(cont As ComboBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    cont.AddItem " "
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!tab_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub



Public Sub LLENA_ARTI_PRO(wcodclie As Currency)
Dim WARTI As rdoQuery
Dim wllave_arti As rdoResultset

pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ? AND ART_CODCLIE = ? AND ART_CALIDAD = 1 ORDER BY ART_NOMBRE "
Set WARTI = CN.CreateQuery("", pub_cadena)
WARTI.rdoParameters(0) = " "
WARTI.rdoParameters(1) = 0
Set wllave_arti = WARTI.OpenResultset(rdOpenKeyset, rdConcurValues)
WARTI(0) = LK_CODCIA
WARTI(1) = wcodclie
wllave_arti.Requery
lisarti.Clear
If wllave_arti.EOF Then
  lisarti.Clear
  Exit Sub
End If
Do Until wllave_arti.EOF
  lisarti.AddItem wllave_arti!art_alterno & " " & wllave_arti!art_nombre & String(120, " ") & wllave_arti!ART_KEY
wllave_arti.MoveNext
Loop


End Sub
Public Sub LLENA_ARTI(wfami As Integer, WSUBFAMI As Integer)
Dim WARTI As rdoQuery
Dim wllave_arti As rdoResultset
Dim wvalor As Integer
If wfami <> 0 Then
  pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ? AND ART_FAMILIA = ? AND ART_CALIDAD = 1 and ART_FLAG_STOCK = 'M' ORDER BY ART_ALTERNO "
  wvalor = wfami
ElseIf WSUBFAMI <> 0 Then
  pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ? AND ART_SUBFAM = ? AND ART_CALIDAD = 1 AND ART_FLAG_STOCK = 'M' ORDER BY ART_ALTERNO "
  wvalor = WSUBFAMI
Else
  Exit Sub
End If
Set WARTI = CN.CreateQuery("", pub_cadena)
WARTI(0) = " "
WARTI(1) = 0
Set wllave_arti = WARTI.OpenResultset(rdOpenKeyset, rdConcurValues)
WARTI(0) = LK_CODCIA
WARTI(1) = wvalor
wllave_arti.Requery
lisarti.Clear
If wllave_arti.EOF Then
  lisarti.Clear
  Exit Sub
End If
Do Until wllave_arti.EOF
  lisarti.AddItem wllave_arti!art_alterno & " " & wllave_arti!art_nombre & String(120, " ") & wllave_arti!ART_KEY
wllave_arti.MoveNext
Loop


End Sub

Private Sub tpromedio_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
'If KeyAscii = 13 Then
' Azul txtagencia, txtagencia
'End If
End Sub

Private Sub tserie_GotFocus()
fraListar.Visible = False
End Sub

Private Sub tserie_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
txtdoc.SetFocus

End Sub

Private Sub txtagencia_GotFocus()
    temporal = Trim(txtagencia.Text)
End Sub

Private Sub txtagencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      txtContacto.SetFocus
    End If
End Sub

Private Sub txtagencia_KeyUp(KeyCode As Integer, Shift As Integer)

'If KeyCode = 45 Then
'PUB_TIPREG = -10
'PUB_CODCIA = LK_CODCIA
'Load FrmDatArti
'FrmDatArti.Caption = "Mantenimiento de Transportistas"
'FrmDatArti.Show 1
'PRO_V.Requery
'txtagencia.Clear
'Do Until PRO_V.EOF
'    txtagencia.AddItem Trim(PRO_V!TRN_NOMBRE) & String(80, " ") & PRO_V!TRN_KEY
'    PRO_V.MoveNext
'Loop
'txtagencia.SetFocus
'DoEvents
'
'End If

End Sub

Private Sub txtagencia_LostFocus()
On Error GoTo SALE
If Trim(temporal) = Trim(txtagencia.Text) Then Exit Sub
par_llave.Edit
par_llave!PAR_AGE_EMP = Trim(txtagencia.Text)
par_llave.Update
par_llave.Requery
Exit Sub
SALE:
End Sub

Private Sub txtcontacto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   fpago.SetFocus
End If

End Sub

Private Sub txtdias_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  Azul tpromedio, tpromedio
End If
End Sub

Private Sub txtdoc_GotFocus()
    fraListar.Visible = False
End Sub

Public Function PROCESA_CANTIDAD(WCODART As Currency, WDIAS As Integer) As Currency
Dim fx As Integer
Dim fecha1
Dim fecha2
Dim wdividir As Integer
Dim WFLAG  As String * 1
Dim WARTI As rdoQuery
Dim wllave_arti As rdoResultset
Dim wrango As Integer
Dim SUM_cantidad As Currency
Dim WQ_cANTIDAD As Currency
PSLOC_WARTI(0) = 0
PSLOC_WARTI(1) = 0
PSLOC_WARTI(2) = 0

For fx = 0 To LISCIA.ListCount - 1
 LISCIA.ListIndex = fx
 If LISCIA.Selected(fx) Then
  If Val(PSLOC_WARTI(0)) = 0 Then
    PSLOC_WARTI(0) = Left(LISCIA.Text, 2)
    GoTo dale
  End If
  If Val(PSLOC_WARTI(1)) = 0 Then
    PSLOC_WARTI(1) = Left(LISCIA.Text, 2)
    GoTo dale
  End If
  If Val(PSLOC_WARTI(2)) = 0 Then
    PSLOC_WARTI(2) = Left(LISCIA.Text, 2)
    GoTo dale
  End If
 End If
dale:
Next fx


PSLOC_WARTI(3) = WCODART
wdividir = 0
wrango = Val(tpromedio.Text) * -1
PROCESA_CANTIDAD = 0
fecha1 = DateAdd("d", wrango, LK_FECHA_DIA)
fecha2 = DateAdd("d", WDIAS, fecha1)
WFLAG = "S"
Do Until WFLAG = "N"
    If fecha1 > LK_FECHA_DIA Then
       Exit Do
    End If
    PSLOC_WARTI(4) = fecha1
    PSLOC_WARTI(5) = fecha2
    llave_sum_arti.Requery
    If Not llave_sum_arti.EOF Then
      If Nulo_Valor0(llave_sum_arti!Cantidad) <> 0 Then
        wdividir = wdividir + 1
        SUM_cantidad = SUM_cantidad + Nulo_Valor0(llave_sum_arti!Cantidad)
      End If
    End If
    
    fecha1 = DateAdd("d", 1, fecha2)
    fecha2 = DateAdd("d", WDIAS, fecha1)
Loop
If SUM_cantidad <> 0 Then
WQ_cANTIDAD = Format((SUM_cantidad / wdividir), "000000")
Else
  SUM_cantidad = 0
End If
PROCESA_CANTIDAD = WQ_cANTIDAD

End Function

Private Sub txtdoc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
txtagencia.SetFocus
End Sub

Private Sub txtdoc_LostFocus()
Screen.MousePointer = 11
DoEvents
Dim indice As Integer
barra.Visible = True
DoEvents
exito = True
fila = 1
SUM_D = 0
SUM_H = 0
CABE_MAN
'LIMPIA_DATOS
GoSub ACT1

fila = 0
barra.Visible = False
grid_fac.SetFocus
GoTo fin

ACT1:
PSTEMP_LLAVE.rdoParameters(0) = LK_CODCIA
PSTEMP_LLAVE.rdoParameters(1) = Val(tserie.Text)
PSTEMP_LLAVE.rdoParameters(2) = Val(txtdoc.Text)
temp_llave.Requery
If temp_llave.EOF Then
 MsgBox "No Existe Docuemnto...", 48, Pub_Titulo
 GoTo borra
End If

barra.Min = 0
barra.max = temp_llave.RowCount
barra.Value = 0

fila = 2
indice = 0
For fila = 0 To codpro.ListCount - 1
''Do Until indice = codpro.ListCount - 1
    codpro.ListIndex = fila
    If Val(Right(codpro.Text, 9)) = temp_llave!PED_CODCLIE Then Exit For
    ''indice = indice + 1
Next fila
''Loop

'indice = 0
'Do Until indice = txtagencia.ListCount - 1
'    txtagencia.ListIndex = indice
    txtagencia.Text = UCase(Nulo_Valors(temp_llave!PED_OFERTA))
'    indice = indice + 1
'Loop
'txtagencia.Text = Nulo_Valors(temp_llave!PED_OFERTA)
txtContacto.Text = Trim(Nulo_Valors(temp_llave!PED_CONTACTO))
Combo1.Text = Trim(Nulo_Valors(temp_llave!PED_NOMCLIE))
On Error GoTo PASES
For fila = 0 To fpago.ListCount - 1
 fpago.ListIndex = fila
 If Trim(UCase(Trim(Left(fpago.Text, 30)))) = Trim(UCase(Left(temp_llave!PED_FORMA, 30))) Then
   Exit For
 End If
Next fila
If Nulo_Valors(temp_llave!PED_MONEDA) = "S" Then
 moneda.ListIndex = 0
Else
 moneda.ListIndex = 1
End If
On Error GoTo 0
fecha.Caption = Format(temp_llave!PED_FECHA, "dd/mm/yyyy")
fila = 0
fila = 2

Do Until temp_llave.EOF
   grid_fac.Rows = fila + 2
   grid_fac.TextMatrix(fila, 2) = Format(temp_llave!PED_CANTIDAD, "0.00")
   grid_fac.TextMatrix(fila, 4) = Format(temp_llave!PED_PRECIO, "0.0000")
   grid_fac.TextMatrix(fila, 3) = temp_llave!PED_UNIDAD
   grid_fac.TextMatrix(fila, 9) = temp_llave!PED_CODART
   grid_fac.TextMatrix(fila, 5) = temp_llave!PED_SUBTOTAL
   grid_fac.TextMatrix(fila, 19) = 5
   SQ_OPER = 2
   PUB_CODART = temp_llave!PED_CODART
   pu_codcia = LK_CODCIA
   LEER_PRE_LLAVE
   grid_fac.TextMatrix(fila, 18) = 0
   Do Until pre_mayor.EOF
    If Val(pre_mayor!PRE_EQUIV) = Val(temp_llave!PED_EQUIV) Then
        grid_fac.TextMatrix(fila, 18) = Val(pre_mayor!pre_PESO)
        Exit Do
    End If
   pre_mayor.MoveNext
   Loop
   
   SQ_OPER = 1
   PUB_KEY = temp_llave!PED_CODART
   pu_codcia = LK_CODCIA
   LEER_ART_LLAVE
   grid_fac.TextMatrix(fila, 0) = "'" & art_LLAVE!art_nombre
   grid_fac.TextMatrix(fila, 1) = art_LLAVE!art_alterno
   grid_fac.TextMatrix(fila, 20) = temp_llave("Ped_DESCTO") 'PORCENTAJE
   grid_fac.TextMatrix(fila, 19) = temp_llave("Ped_DESCTO_PRE") 'VALOR
   If fila = 2 Then
   End If
   fila = fila + 1
   barra.Value = fila - 2
   temp_llave.MoveNext
Loop
suma_grid
f1.Enabled = True
ESTADO.Enabled = True

Return

fin:
Screen.MousePointer = 0
Exit Sub
borra:
cancelar_Click
Screen.MousePointer = 0
Exit Sub
PASES:
Resume Next


End Sub

Private Sub UNIDAD_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode <> 45 Then Exit Sub


SQ_OPER = 1
pu_codcia = LK_CODCIA
PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 9))
PUB_SECUEN = Val(Right(UNIDAD.Text, 4))
LEER_PRE_LLAVE
grid_fac.TextMatrix(grid_fac.Row, 12) = pre_llave!pre_secuencia

grid_fac.TextMatrix(grid_fac.Row, 3) = InputBox("Unidad: ")

'grid_fac.TextMatrix(grid_fac.Row, 4) = Format(Val(grid_fac.TextMatrix(grid_fac.Row, 11)) / Val(grid_fac.TextMatrix(grid_fac.Row, 17)), "0.00")
'grid_fac.TextMatrix(grid_fac.Row, 17) = pre_llave!pre_equiv
'grid_fac.TextMatrix(grid_fac.Row, 6) = Format(grid_fac.TextMatrix(grid_fac.Row, 14) / pre_llave!pre_equiv, "#######")
'grid_fac.TextMatrix(grid_fac.Row, 7) = Format(grid_fac.TextMatrix(grid_fac.Row, 15) / pre_llave!pre_equiv, "#######")
'grid_fac.TextMatrix(grid_fac.Row, 8) = Format(grid_fac.TextMatrix(grid_fac.Row, 16) / pre_llave!pre_equiv, "#######")
    


UNIDAD.Visible = False
'suma_grid
grid_fac.COL = 4
grid_fac.SetFocus


End Sub

Private Sub UNIDAD_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SQ_OPER = 1
pu_codcia = LK_CODCIA
PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 9))
PUB_SECUEN = Val(Right(UNIDAD.Text, 4))
LEER_PRE_LLAVE
grid_fac.TextMatrix(grid_fac.Row, 12) = pre_llave!pre_secuencia
grid_fac.TextMatrix(grid_fac.Row, 3) = Trim(Left(UNIDAD.Text, 12))
grid_fac.TextMatrix(grid_fac.Row, 4) = Format(Val(grid_fac.TextMatrix(grid_fac.Row, 11)) / Val(grid_fac.TextMatrix(grid_fac.Row, 17)), "0.00")
grid_fac.TextMatrix(grid_fac.Row, 17) = pre_llave!PRE_EQUIV
grid_fac.TextMatrix(grid_fac.Row, 6) = Format(grid_fac.TextMatrix(grid_fac.Row, 14) / pre_llave!PRE_EQUIV, "#######")
grid_fac.TextMatrix(grid_fac.Row, 7) = Format(grid_fac.TextMatrix(grid_fac.Row, 15) / pre_llave!PRE_EQUIV, "#######")
grid_fac.TextMatrix(grid_fac.Row, 8) = Format(grid_fac.TextMatrix(grid_fac.Row, 16) / pre_llave!PRE_EQUIV, "#######")
    


UNIDAD.Visible = False
suma_grid
grid_fac.COL = 4
grid_fac.SetFocus

End Sub
