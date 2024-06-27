VERSION 5.00
Begin VB.Form frmConfigRepo 
   Caption         =   "Configuración de Reportes"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   1275
   ClientWidth     =   11595
   Icon            =   "frmConfigRepo.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12797.87
   ScaleMode       =   0  'User
   ScaleWidth      =   11595
   Begin VB.Frame fraingreso 
      Caption         =   "Configuración de Reporte"
      Height          =   5895
      Left            =   480
      TabIndex        =   43
      Top             =   360
      Visible         =   0   'False
      Width           =   10575
      Begin VB.CheckBox CHEFLAG2 
         Height          =   255
         Left            =   7680
         TabIndex        =   108
         Top             =   3120
         Width           =   375
      End
      Begin VB.CheckBox cheprofa 
         Caption         =   "Fechas Proceso FACART"
         Height          =   255
         Left            =   7680
         TabIndex        =   107
         ToolTipText     =   "Crear Formula : DIA"
         Top             =   2520
         Width           =   2175
      End
      Begin VB.CheckBox CHEMOFAR 
         Caption         =   "Moneda  (FACART)"
         Height          =   255
         Left            =   7680
         TabIndex        =   105
         ToolTipText     =   "Crear Formula : DIA"
         Top             =   3960
         Width           =   2175
      End
      Begin VB.CheckBox CHEMOCAR 
         Caption         =   "Moneda  (CARTERA)"
         Height          =   255
         Left            =   7680
         TabIndex        =   104
         ToolTipText     =   "Crear Formula : DIA"
         Top             =   3720
         Width           =   2175
      End
      Begin VB.CheckBox chemon 
         Caption         =   "Moneda Caja (ALLOG)"
         Height          =   255
         Left            =   7680
         TabIndex        =   103
         ToolTipText     =   "Crear Formula : DIA"
         Top             =   3480
         Width           =   2175
      End
      Begin VB.CheckBox chetipo 
         Caption         =   "Configurar Tipo de Cliente"
         Height          =   255
         Left            =   5160
         TabIndex        =   102
         ToolTipText     =   "Crear Formula : DIA"
         Top             =   3960
         Width           =   2175
      End
      Begin VB.CheckBox CHEARM 
         Caption         =   "Listar Compañias - ARTICULO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7440
         TabIndex        =   101
         Top             =   2160
         Width           =   3375
      End
      Begin VB.CheckBox chegiro 
         Caption         =   "ALL_FECHA_CAN (Fecha Giro )"
         Height          =   255
         Left            =   7440
         TabIndex        =   100
         Top             =   1800
         Width           =   3135
      End
      Begin VB.CheckBox CHEFLAG 
         Height          =   255
         Left            =   7680
         TabIndex        =   98
         Top             =   2880
         Width           =   375
      End
      Begin VB.CheckBox checiaall 
         Caption         =   "Listar Compañias para ALLOG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4440
         TabIndex        =   97
         Top             =   1080
         Width           =   3375
      End
      Begin VB.CheckBox chefechapro 
         Caption         =   "ALL_FECHA_PRO (Fecha Contable )"
         Height          =   255
         Left            =   7440
         TabIndex        =   96
         Top             =   1440
         Width           =   3135
      End
      Begin VB.CheckBox checiacar 
         Caption         =   "Listar Compañias para Ctas,Ctes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4440
         TabIndex        =   94
         Top             =   840
         Width           =   3375
      End
      Begin VB.CheckBox cheped 
         Caption         =   "Fechas Pedidos"
         Height          =   255
         Left            =   5160
         TabIndex        =   91
         ToolTipText     =   "Crear Formula : DIA"
         Top             =   3720
         Width           =   2175
      End
      Begin VB.CheckBox chestock 
         Caption         =   "Flag / Stock"
         Height          =   255
         Left            =   840
         TabIndex        =   90
         ToolTipText     =   "Crear Formula : DIA"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.CheckBox checcm 
         Caption         =   "Cia de Comaest"
         Height          =   255
         Left            =   5160
         TabIndex        =   87
         ToolTipText     =   "Crear Formula : DIA"
         Top             =   3480
         Width           =   2175
      End
      Begin VB.CheckBox chenat 
         Caption         =   "Report. Natur."
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
         Height          =   255
         Left            =   8640
         TabIndex        =   86
         ToolTipText     =   "Crear Formula : DIA"
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox chesincia 
         Caption         =   "Listar Compañias para Productos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4440
         TabIndex        =   83
         Top             =   600
         Width           =   3375
      End
      Begin VB.CheckBox cheusu 
         Caption         =   "Filtrar x Cod. Usuario "
         Height          =   255
         Left            =   5160
         TabIndex        =   82
         ToolTipText     =   "Crear Formula : DIA"
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CheckBox checar 
         Caption         =   "Fecha Ingreso Cartera"
         Height          =   255
         Left            =   5160
         TabIndex        =   81
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CheckBox chestockR 
         Caption         =   "Solo Stock x Regla"
         Height          =   195
         Left            =   5160
         TabIndex        =   80
         Top             =   5040
         Width           =   2655
      End
      Begin VB.CheckBox cheestado 
         Caption         =   "Opcion Estado de Cliente"
         Height          =   195
         Left            =   5160
         TabIndex        =   79
         Top             =   3240
         Width           =   2175
      End
      Begin VB.CheckBox chesituacion 
         Caption         =   "Fechas Cheques"
         Height          =   255
         Left            =   2400
         TabIndex        =   77
         ToolTipText     =   "Crear Formula : DIA"
         Top             =   2880
         Width           =   1935
      End
      Begin VB.CheckBox Chetipdoc 
         Caption         =   "UtilizarTip. Doc.(CARTRA)"
         Height          =   255
         Left            =   2880
         TabIndex        =   76
         ToolTipText     =   "Crear Formula : DIA"
         Top             =   2520
         Width           =   2175
      End
      Begin VB.CheckBox chefecha 
         Caption         =   "Pasar Fecha del Sistema"
         Height          =   255
         Left            =   240
         TabIndex        =   75
         ToolTipText     =   "Crear Formula : DIA"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox checolu 
         Caption         =   "Utilizar Fecha como Columnas"
         Height          =   255
         Left            =   2400
         TabIndex        =   74
         ToolTipText     =   "Crear Formula : DIA"
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CheckBox chepto 
         Caption         =   "Reporte para Pto. de Ventas."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   72
         ToolTipText     =   "Crear Formula : CIA"
         Top             =   240
         Width           =   2895
      End
      Begin VB.CheckBox cheallog 
         Caption         =   "Fechas (ALLOG )"
         Height          =   255
         Left            =   2880
         TabIndex        =   71
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CheckBox cherango 
         Caption         =   "Mostrar Rango de Fechas"
         Height          =   255
         Left            =   2400
         TabIndex        =   69
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CheckBox checaa 
         Caption         =   "Fechas( Amortizado.CARACU)"
         Height          =   375
         Left            =   5160
         TabIndex        =   68
         Top             =   2400
         Width           =   2535
      End
      Begin VB.CheckBox checlipro 
         Caption         =   "Opcion Cliente/Proveedor"
         Height          =   195
         Left            =   5160
         TabIndex        =   67
         Top             =   2880
         Width           =   2175
      End
      Begin VB.CheckBox chefechas1 
         Caption         =   "Fechas ( New Vcto. CARTERA)"
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   2520
         Width           =   2775
      End
      Begin VB.CheckBox checia 
         Caption         =   "Mostrar Compañia"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         ToolTipText     =   "Crear Formula : CIA"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox chedia 
         Caption         =   "Mostrar Fecha del sistema"
         Height          =   255
         Left            =   5160
         TabIndex        =   64
         ToolTipText     =   "Crear Formula : DIA"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox chebancos 
         Caption         =   "Codigo de Bancos"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CheckBox chevendedor 
         Caption         =   "Codigo de Vendedor"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CheckBox chefechas 
         Caption         =   "Fechas (FACART )"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   285
         Left            =   3960
         TabIndex        =   48
         Top             =   5520
         Width           =   1455
      End
      Begin VB.CommandButton cmdgrabar 
         Caption         =   "&Grabar"
         Height          =   285
         Left            =   2280
         TabIndex        =   47
         Top             =   5520
         Width           =   1215
      End
      Begin VB.TextBox txtformulario 
         Height          =   285
         Left            =   1680
         MaxLength       =   25
         TabIndex        =   46
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtarchivo 
         Height          =   285
         Left            =   120
         MaxLength       =   25
         TabIndex        =   45
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtnombre 
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   44
         Top             =   480
         Width           =   3135
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   5160
         TabIndex        =   55
         Top             =   4200
         Width           =   3015
         Begin VB.OptionButton opcp 
            Caption         =   "Utilizar codigo Cliente"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   57
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton opcp 
            Caption         =   "Utilizar codigo Proveedor"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   56
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1815
         Left            =   120
         TabIndex        =   58
         Top             =   3480
         Width           =   4815
         Begin VB.CheckBox chearti 
            Caption         =   "Codigo de Producto"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   2640
            TabIndex        =   106
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CheckBox checal 
            Caption         =   "Calidades del Prod."
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   2640
            TabIndex        =   93
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CheckBox cheplancha 
            Caption         =   "Planchas"
            Height          =   255
            Left            =   2640
            TabIndex        =   92
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CheckBox chesubgru 
            Caption         =   "Sub Grupo"
            Height          =   255
            Left            =   2640
            TabIndex        =   89
            Top             =   120
            Width           =   1575
         End
         Begin VB.CheckBox chelinea 
            Caption         =   "Lineas"
            Height          =   255
            Left            =   2640
            TabIndex        =   88
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox chenumero 
            Caption         =   "Utilizar Numero"
            Height          =   255
            Left            =   2640
            TabIndex        =   85
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chemarca 
            Caption         =   "Utilizar Marca"
            Height          =   255
            Left            =   2640
            TabIndex        =   84
            Top             =   840
            Width           =   1575
         End
         Begin VB.CheckBox chemoneda 
            Caption         =   "Mostrar Opcion de Monedas"
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   960
            Width           =   2415
         End
         Begin VB.CheckBox chetablas 
            Caption         =   "Aplicar a ""TABLAS"""
            Height          =   195
            Left            =   720
            TabIndex        =   70
            ToolTipText     =   "Crear Formula : CIA"
            Top             =   1560
            Width           =   2295
         End
         Begin VB.OptionButton oparti 
            Caption         =   "Utilizar Sub. Familias  Art."
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   63
            Top             =   700
            Width           =   2175
         End
         Begin VB.OptionButton oparti 
            Caption         =   "Utilizar Familias de Articulo"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   60
            Top             =   420
            Width           =   2175
         End
         Begin VB.OptionButton oparti 
            Caption         =   "Utilizar codigo Articulo"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   59
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.CheckBox Chezonas 
         Caption         =   "Zonas y Sub-Zonas deClientes"
         Height          =   195
         Left            =   2400
         TabIndex        =   62
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label LBLFLAG2 
         Caption         =   "Flag"
         Height          =   255
         Left            =   8040
         TabIndex        =   109
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lblflag 
         Caption         =   "Flag"
         Height          =   255
         Left            =   8040
         TabIndex        =   99
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Filtro por Fechas"
         Height          =   255
         Left            =   7920
         TabIndex        =   95
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "(solo marcar para reportes que usen la tabla ""ARTI"" )"
         Height          =   255
         Left            =   6600
         TabIndex        =   73
         Top             =   240
         Width           =   4455
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filtro de Reportes ( Crystal Report )"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   61
         Top             =   1920
         Width           =   2460
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre del Formulario :"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   51
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre del Archivo :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   50
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Titulo de Reporte :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fraconf 
      Caption         =   "8"
      Height          =   2895
      Index           =   6
      Left            =   8520
      TabIndex        =   17
      Top             =   2880
      Width           =   2895
      Begin VB.CheckBox cherep 
         Caption         =   "Activo"
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   32
         Tag             =   "8"
         Top             =   600
         Width           =   855
      End
      Begin VB.ListBox listarep8 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   7
         Tag             =   "8"
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label lblmensa 
         Caption         =   "Un Momento .."
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   40
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label txtrep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   480
         TabIndex        =   25
         Tag             =   "8"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fraconf 
      Caption         =   "7"
      Height          =   2895
      Index           =   4
      Left            =   5640
      TabIndex        =   15
      Top             =   2880
      Width           =   2895
      Begin VB.CheckBox cherep 
         Caption         =   "Activo"
         Height          =   195
         Index           =   6
         Left            =   480
         TabIndex        =   31
         Tag             =   "7"
         Top             =   600
         Width           =   855
      End
      Begin VB.ListBox listarep7 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   6
         Tag             =   "7"
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label lblmensa 
         Caption         =   "Un Momento .."
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   39
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label txtrep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   480
         TabIndex        =   24
         Tag             =   "7"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fraconf 
      Caption         =   "6"
      Height          =   2895
      Index           =   2
      Left            =   2760
      TabIndex        =   13
      Top             =   2880
      Width           =   2895
      Begin VB.CheckBox cherep 
         Caption         =   "Activo"
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   30
         Tag             =   "6"
         Top             =   600
         Width           =   855
      End
      Begin VB.ListBox listarep6 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   5
         Tag             =   "6"
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label lblmensa 
         Caption         =   "Un Momento .."
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   38
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label txtrep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   360
         TabIndex        =   23
         Tag             =   "6"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdagregar 
      Caption         =   "Agregar Reportes"
      Height          =   375
      Left            =   5040
      TabIndex        =   42
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "Ce&rrar"
      Height          =   375
      Left            =   7800
      TabIndex        =   8
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Frame fraconf 
      Caption         =   "4"
      Height          =   2775
      Index           =   5
      Left            =   8520
      TabIndex        =   16
      Top             =   0
      Width           =   2895
      Begin VB.CheckBox cherep 
         Caption         =   "Activo"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   28
         Tag             =   "4"
         Top             =   600
         Width           =   855
      End
      Begin VB.ListBox listarep4 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   3
         Tag             =   "4"
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label lblmensa 
         Caption         =   "Un Momento .."
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   36
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label txtrep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   480
         TabIndex        =   21
         Tag             =   "4"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fraconf 
      Caption         =   "3"
      Height          =   2775
      Index           =   3
      Left            =   5640
      TabIndex        =   14
      Top             =   0
      Width           =   2895
      Begin VB.CheckBox cherep 
         Caption         =   "Activo"
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   27
         Tag             =   "3"
         Top             =   600
         Width           =   855
      End
      Begin VB.ListBox listarep3 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   2
         Tag             =   "3"
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label lblmensa 
         Caption         =   "Un Momento .."
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   35
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label txtrep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   480
         TabIndex        =   20
         Tag             =   "3"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fraconf 
      Caption         =   "2"
      Height          =   2775
      Index           =   0
      Left            =   2760
      TabIndex        =   12
      Top             =   0
      Width           =   2895
      Begin VB.CheckBox cherep 
         Caption         =   "Activo"
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   26
         Tag             =   "2"
         Top             =   600
         Width           =   855
      End
      Begin VB.ListBox listarep2 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   1
         Tag             =   "2"
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label lblmensa 
         Caption         =   "Un Momento .."
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   34
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label txtrep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   19
         Tag             =   "2"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fraconf 
      Caption         =   "1"
      Height          =   2775
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   2775
      Begin VB.CheckBox cherep 
         Caption         =   "Activo"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   10
         Tag             =   "1"
         Top             =   600
         Width           =   855
      End
      Begin VB.ListBox listarep1 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         ItemData        =   "frmConfigRepo.frx":0442
         Left            =   120
         List            =   "frmConfigRepo.frx":044F
         Style           =   1  'Checkbox
         TabIndex        =   0
         Tag             =   "1"
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label lblmensa 
         Caption         =   "Un Momento .."
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   33
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label txtrep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   18
         Tag             =   "1"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fraconf 
      Caption         =   "5"
      Height          =   2895
      Index           =   8
      Left            =   0
      TabIndex        =   11
      Top             =   2880
      Width           =   2775
      Begin VB.CheckBox cherep 
         Caption         =   "Activo"
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   29
         Tag             =   "5"
         Top             =   600
         Width           =   855
      End
      Begin VB.ListBox listarep5 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   4
         Tag             =   "5"
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label lblmensa 
         Caption         =   "Un Momento .."
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   37
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label txtrep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   360
         TabIndex        =   22
         Tag             =   "5"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Label lblmensaje 
      Height          =   255
      Left            =   240
      TabIndex        =   41
      Top             =   5880
      Width           =   4455
   End
End
Attribute VB_Name = "frmConfigRepo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ws_nombre As String
Dim ws_nombre2 As String
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset
Dim PS_REP02 As rdoQuery
Dim llave_rep02 As rdoResultset
Dim PS_REP03 As rdoQuery
Dim llave_rep03 As rdoResultset
Dim REP_CODTRA As Currency
Dim WS_FLAG As String * 1
Dim loc_flag_cambio  As String * 1
Dim wsINI1 As Currency
Dim wsFIN1 As Currency
Dim wsINI2 As Currency
Dim wsFIN2 As Currency
Dim wsINI3 As Currency
Dim wsFIN3 As Currency
Dim wsINI4 As Currency
Dim wsFIN4 As Currency
Dim wsINI5 As Currency
Dim wsFIN5 As Currency
Dim wsINI6 As Currency
Dim wsFIN6 As Currency
Dim wsINI7 As Currency
Dim wsFIN7 As Currency
Dim wsINI8 As Currency
Dim wsFIN8 As Currency

Private Sub chenat_Click()
If chenat.Value = 0 Then
 wsc = UCase(InputBox("Clave de aceeso: ", "Acceso Limitado", "**********"))
 If wsc <> PUB_CLAVE Then
    chenat.Value = 1
    Exit Sub
 End If
End If

End Sub

Private Sub cherep_Click(Index As Integer)
If Trim(cherep(Index).Tag) = "" Or loc_flag_cambio = "X" Then
 Exit Sub
End If
loc_flag_cambio = "A"
SQ_OPER = 1
PUB_TIPREG = 99
PUB_NUMTAB = Val(cherep(Index).Tag)
PUB_CODCIA = "00"
LEER_TAB_LLAVE
If tab_llave.EOF Then
 Exit Sub
End If
  tab_llave.Edit
  tab_llave!TAB_NOMCORTO = cherep(Index).Value
  tab_llave.Update
Screen.MousePointer = 0


End Sub

Private Sub cherep_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmensaje.Caption = "Activa o desactiva el bloque de reporte..."
End Sub

Private Sub cmdagregar_Click()
WS_FLAG = ""
LIMPIA_DATOS
fraingreso.Visible = True
txtnombre.SetFocus

End Sub

Private Sub cmdagregar_GotFocus()
If fraingreso.Visible Then fraingreso.Visible = False
End Sub

Private Sub cmdagregar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmensaje.Caption = ""
End Sub

Private Sub cmdcancelar_Click()
SQ_OPER = 0
WS_FLAG = ""
PUB_CODTRA = 0
LIMPIA_DATOS
fraingreso.Visible = False
cmdagregar.SetFocus
End Sub

Private Sub cmdcerrar_Click()
Unload frmConfigRepo
End Sub

Private Sub cmdcerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmensaje.Caption = ""
End Sub

Private Sub cmdgrabar_Click()
Dim wscodtra As Currency
Dim wvalor
Dim POSI As Integer
cmdgrabar.Enabled = False
If txtnombre.text = "" Then
 MsgBox "Ingrese Descripción...", 48, Pub_Titulo
 Azul txtnombre, txtnombre
 cmdgrabar.Enabled = True
 Exit Sub
End If
If txtarchivo.text = "" Then
 MsgBox "Ingrese Archivo...", 48, Pub_Titulo
 Azul txtarchivo, txtarchivo
 cmdgrabar.Enabled = True
 Exit Sub
End If
If txtformulario.text = "" Then
 MsgBox "Ingrese Formulario...", 48, Pub_Titulo
 Azul txtformulario, txtformulario
 cmdgrabar.Enabled = True
 Exit Sub
End If
If WS_FLAG = "E" Then
 GoTo EDITAR
End If

wvalor = InputBox("Ingrese  Numero de tabla(colocación)  : (valor entre 1 y 8 )", Pub_Titulo)
If wvalor = "" Then
  cmdgrabar.Enabled = True
  Exit Sub
End If
If Val(wvalor) < 1 And Val(wvalor) > 8 Then
   Screen.MousePointer = 0
   MsgBox "Deber ser entre 1 y 8 ... ", 48, Pub_Titulo
   cmdgrabar.Enabled = True
   Exit Sub
End If
POSI = wvalor
wsarchivo = Left(txtarchivo.text, 25)
wsformulario = Left(txtformulario.text, 25)
wsNOMBRE = Trim(txtnombre.text)
wscodtra = NEW_TRA_KEY(POSI)
If wscodtra = 0 Then
  Screen.MousePointer = 0
  MsgBox "Fuera de Lugar ...Intente nuevamente ", 48, Pub_Titulo
  cmdgrabar.Enabled = True
  Exit Sub
End If
llave_rep02.AddNew
For fila = 0 To llave_rep02.rdoColumns.Count - 1
 llave_rep02.rdoColumns(fila) = 0
Next fila
llave_rep02!TRA_KEY = wscodtra
llave_rep02!tra_descripcion = wsNOMBRE
llave_rep02!TRA_L1 = wsarchivo
llave_rep02!TRA_L2 = wsformulario
llave_rep02!TRA_FLAG_ACTIVO = "A"
llave_rep02!TRA_FLAG_EXT = "A"
llave_rep02!tra_ACT10 = cherango.Value
llave_rep02!tra_con1 = chefechas.Value
llave_rep02!TRA_CON2 = chebancos.Value
llave_rep02!TRA_CON3 = chevendedor.Value
llave_rep02!tra_rep1 = chepto.Value
llave_rep02!TRA_CON4 = 0
llave_rep02!TRA_CON5 = 0
If opcp(0).Value Then
   llave_rep02!TRA_CON4 = 1
End If
If opcp(1).Value Then
   llave_rep02!TRA_CON5 = 1
End If
llave_rep02!tra_con6 = 0
llave_rep02!tra_con7 = 0
llave_rep02!tra_con9 = 0
If oparti(0).Value Then
 llave_rep02!tra_con6 = 1
End If
If oparti(1).Value Then
 llave_rep02!tra_con7 = 1
End If
If oparti(2).Value Then
 llave_rep02!tra_con9 = 1
End If
llave_rep02!TRA_CON8 = Chezonas.Value
llave_rep02!tra_con10 = chefechas1.Value
llave_rep02!TRA_CON11 = checlipro.Value
llave_rep02!tra_con12 = checaa.Value
llave_rep02!TRA_CON13 = chetablas.Value
llave_rep02!TRA_CON14 = cheallog.Value
llave_rep02!TRA_CON15 = checolu.Value
llave_rep02!tra_ACT1 = checia.Value
llave_rep02!tra_ACT2 = chedia.Value
llave_rep02!TRA_ACT3 = chefecha.Value
llave_rep02!TRA_ACT4 = Chetipdoc.Value
llave_rep02!TRA_ACT5 = chesituacion.Value
llave_rep02!TRA_ACT6 = chemoneda.Value
llave_rep02!TRA_ACT7 = cheestado.Value
llave_rep02!TRA_ACT8 = checar.Value
llave_rep02!TRA_ACT9 = cheusu.Value
llave_rep02!TRA_ACT11 = chesincia.Value
llave_rep02!TRA_ACT12 = chenumero.Value
llave_rep02!TRA_ACT13 = chemarca.Value
llave_rep02!TRA_ACT14 = chenat.Value
llave_rep02!TRA_ACT15 = chelinea.Value
llave_rep02!TRA_S1 = checcm.Value
llave_rep02!TRA_S2 = chesubgru.Value
llave_rep02!TRA_S3 = chestock.Value
llave_rep02!TRA_S4 = cheped.Value
llave_rep02!TRA_S5 = cheplancha.Value
llave_rep02!TRA_S6 = checiacar.Value
llave_rep02!TRA_S7 = checal.Value
llave_rep02!TRA_S8 = chefechapro.Value
llave_rep02!TRA_S9 = checiaall.Value
llave_rep02!TRA_S10 = CHEFLAG.Value
llave_rep02!TRA_L3 = ws_nombre
llave_rep02!TRA_L4 = ws_nombre2

llave_rep02!TRA_S11 = chegiro.Value
llave_rep02!TRA_S12 = CHEARM.Value
llave_rep02!TRA_S13 = chetipo.Value
llave_rep02!TRA_S14 = chemon.Value
llave_rep02!TRA_s15 = CHEMOCAR.Value
llave_rep02!TRA_C1 = CHEMOFAR.Value
llave_rep02!TRA_C2 = chearti.Value
llave_rep02!TRA_GRU1 = cheprofa.Value
llave_rep02!TRA_GRU2 = CHEFLAG2.Value

llave_rep02.Update
LIMPIA_DATOS
cmdgrabar.Enabled = True
fraingreso.Visible = False
If POSI = 1 Then LLENA_LISTA listarep1, wsINI1, wsFIN1
If POSI = 2 Then LLENA_LISTA listarep2, wsINI2, wsFIN2
If POSI = 3 Then LLENA_LISTA listarep3, wsINI3, wsFIN3
If POSI = 4 Then LLENA_LISTA listarep4, wsINI4, wsFIN4
If POSI = 5 Then LLENA_LISTA listarep5, wsINI5, wsFIN5
If POSI = 6 Then LLENA_LISTA listarep6, wsINI6, wsFIN6
If POSI = 7 Then LLENA_LISTA listarep7, wsINI7, wsFIN7
If POSI = 8 Then LLENA_LISTA listarep8, wsINI8, wsFIN8
cmdagregar.SetFocus
Exit Sub

EDITAR:
PS_REP01(0) = PUB_CODTRA
llave_rep01.Requery
If llave_rep01.EOF Then
   MsgBox "Verificar Key de Transacción ..", 48, Pub_Titulo
  Exit Sub
End If
Screen.MousePointer = 11
wsarchivo = Left(txtarchivo.text, 25)
wsformulario = Left(txtformulario.text, 25)
wsNOMBRE = Trim(txtnombre.text)
llave_rep01.Edit
llave_rep01!tra_descripcion = wsNOMBRE
llave_rep01!TRA_L1 = wsarchivo
llave_rep01!TRA_L2 = wsformulario
llave_rep01!tra_ACT10 = cherango.Value
llave_rep01!tra_con1 = chefechas.Value
llave_rep01!TRA_CON2 = chebancos.Value
llave_rep01!TRA_CON3 = chevendedor.Value
llave_rep01!tra_rep1 = chepto.Value
llave_rep01!TRA_CON4 = 0
llave_rep01!TRA_CON5 = 0
If opcp(0).Value Then
   llave_rep01!TRA_CON4 = 1
End If
If opcp(1).Value Then
   llave_rep01!TRA_CON5 = 1
End If
llave_rep01!tra_con6 = 0
llave_rep01!tra_con7 = 0
llave_rep01!tra_con9 = 0
If oparti(0).Value Then
 llave_rep01!tra_con6 = 1
End If
If oparti(1).Value Then
 llave_rep01!tra_con7 = 1
End If
If oparti(2).Value Then
 llave_rep01!tra_con9 = 1
End If
llave_rep01!TRA_CON8 = Chezonas.Value
llave_rep01!tra_con10 = chefechas1.Value
llave_rep01!TRA_CON11 = checlipro.Value
llave_rep01!tra_ACT1 = checia.Value
llave_rep01!tra_ACT2 = chedia.Value
llave_rep01!tra_con12 = checaa.Value
llave_rep01!TRA_CON13 = chetablas.Value
llave_rep01!TRA_CON14 = cheallog.Value
llave_rep01!TRA_CON15 = checolu.Value
llave_rep01!TRA_ACT3 = chefecha.Value
llave_rep01!TRA_ACT4 = Chetipdoc.Value
llave_rep01!TRA_ACT5 = chesituacion.Value
llave_rep01!TRA_ACT6 = chemoneda.Value
llave_rep01!TRA_ACT7 = cheestado.Value
llave_rep01!TRA_ACT8 = checar.Value
llave_rep01!TRA_ACT9 = cheusu.Value
llave_rep01!TRA_ACT11 = chesincia.Value

llave_rep01!TRA_S2 = chesubgru.Value
llave_rep01!TRA_ACT12 = chenumero.Value
llave_rep01!TRA_ACT15 = chelinea.Value
llave_rep01!TRA_ACT13 = chemarca.Value
llave_rep01!TRA_S5 = cheplancha.Value

llave_rep01!TRA_ACT14 = chenat.Value
llave_rep01!TRA_S1 = checcm.Value
llave_rep01!TRA_S3 = chestock.Value
llave_rep01!TRA_S4 = cheped.Value
llave_rep01!TRA_S7 = checal.Value
llave_rep01!TRA_S6 = checiacar.Value
llave_rep01!TRA_S8 = chefechapro.Value
llave_rep01!TRA_S9 = checiaall.Value
llave_rep01!TRA_S10 = CHEFLAG.Value
llave_rep01!TRA_L3 = ws_nombre
llave_rep01!TRA_L4 = ws_nombre2
llave_rep01!TRA_S11 = chegiro.Value
llave_rep01!TRA_S12 = CHEARM.Value
llave_rep01!TRA_S13 = chetipo.Value
llave_rep01!TRA_S14 = chemon.Value
llave_rep01!TRA_s15 = CHEMOCAR.Value
llave_rep01!TRA_C1 = CHEMOFAR.Value
llave_rep01!TRA_C2 = chearti.Value
llave_rep01!TRA_GRU1 = cheprofa.Value
llave_rep01!TRA_GRU2 = CHEFLAG2.Value







llave_rep01.Update
loc_flag_cambio = "A"
Screen.MousePointer = 0
LIMPIA_DATOS
cmdgrabar.Enabled = True
fraingreso.Visible = False
If SQ_OPER = 1 Then
 LLENA_LISTA listarep1, wsINI1, wsFIN1
 listarep1.SetFocus
ElseIf SQ_OPER = 2 Then
 LLENA_LISTA listarep2, wsINI2, wsFIN2
 listarep1.SetFocus
ElseIf SQ_OPER = 3 Then
 LLENA_LISTA listarep3, wsINI3, wsFIN3
 listarep1.SetFocus
ElseIf SQ_OPER = 4 Then
 LLENA_LISTA listarep4, wsINI4, wsFIN4
 listarep1.SetFocus
ElseIf SQ_OPER = 5 Then
 LLENA_LISTA listarep5, wsINI5, wsFIN5
 listarep1.SetFocus
ElseIf SQ_OPER = 6 Then
 LLENA_LISTA listarep6, wsINI6, wsFIN6
 listarep1.SetFocus
ElseIf SQ_OPER = 7 Then
 LLENA_LISTA listarep7, wsINI7, wsFIN7
 listarep1.SetFocus
ElseIf SQ_OPER = 8 Then
 LLENA_LISTA listarep8, wsINI8, wsFIN8
 listarep1.SetFocus
End If



End Sub

Private Sub Form_Load()
loc_flag_cambio = ""
pub_cadena = "SELECT * FROM TRANSACCION WHERE TRA_KEY = ? ORDER BY TRA_KEY"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM TRANSACCION WHERE TRA_KEY >=?  AND TRA_KEY <= ?  AND TRA_FLAG_ACTIVO = 'A' ORDER BY TRA_DESCRIPCION"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
PS_REP02(1) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM TRANSACCION WHERE TRA_KEY >=?  AND TRA_KEY <= ?  ORDER BY TRA_KEY"
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
PS_REP03(0) = 0
PS_REP03(1) = 0
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenKeyset, rdConcurValues)
wsINI1 = 8100
wsFIN1 = 8299
wsINI2 = 8300
wsFIN2 = 8499
wsINI3 = 8500
wsFIN3 = 8699
wsINI4 = 8700
wsFIN4 = 8899
wsINI5 = 8900
wsFIN5 = 9099
wsINI6 = 9100
wsFIN6 = 9299
wsINI7 = 9300
wsFIN7 = 9499
wsINI8 = 9500
wsFIN8 = 9699

WS_FLAG = "A"
LLENA_LISTA listarep1, wsINI1, wsFIN1
LLENA_LISTA listarep2, wsINI2, wsFIN2
LLENA_LISTA listarep3, wsINI3, wsFIN3
LLENA_LISTA listarep4, wsINI4, wsFIN4
LLENA_LISTA listarep5, wsINI5, wsFIN5
LLENA_LISTA listarep6, wsINI6, wsFIN6
LLENA_LISTA listarep7, wsINI7, wsFIN7
LLENA_LISTA listarep8, wsINI8, wsFIN8
WS_FLAG = ""

SQ_OPER = 1
PUB_CODCIA = "00"
PUB_TIPREG = 340
PUB_NUMTAB = 0
''LEER_TAB_LLAVE
''If Not tab_llave.EOF Then chesubgru.Caption = Trim(tab_llave!TAB_NOMLARGO)
''PUB_NUMTAB = 1
''LEER_TAB_LLAVE
'If Not tab_llave.EOF Then chelinea.Caption = Trim(tab_llave!TAB_NOMLARGO)
PUB_NUMTAB = 2
LEER_TAB_LLAVE
If Not tab_llave.EOF Then chesubgru.Caption = Trim(tab_llave!TAB_NOMLARGO)
PUB_NUMTAB = 3
LEER_TAB_LLAVE
If Not tab_llave.EOF Then chenumero.Caption = Trim(tab_llave!TAB_NOMLARGO)
PUB_NUMTAB = 4
LEER_TAB_LLAVE
If Not tab_llave.EOF Then chelinea.Caption = Trim(tab_llave!TAB_NOMLARGO)
'PUB_NUMTAB = 5
'LEER_TAB_LLAVE
'If Not tab_llave.EOF Then lblart(5).Caption = Trim(tab_llave!TAB_NOMLARGO)


SQ_OPER = 2
PUB_TIPREG = 99
PUB_CODCIA = "00"
LEER_TAB_LLAVE
If tab_mayor.EOF Then
 MsgBox " No existe TIPREG 99 en tablas crearlas con titulos del Menu correspondiente..", 48, Pub_Titulo
 Exit Sub
End If
fila = 0
loc_flag_cambio = "X"
Do Until tab_mayor.EOF
 frmConfigRepo.txtrep(fila).Caption = Trim(tab_mayor!TAB_NOMLARGO)
 frmConfigRepo.txtrep(fila).Tag = Trim(tab_mayor!TAB_NUMTAB)
 frmConfigRepo.cherep(fila).Value = Val(tab_mayor!TAB_NOMCORTO)
 fila = fila + 1
 tab_mayor.MoveNext
Loop
loc_flag_cambio = ""
End Sub

Public Sub LLENA_LISTA(wlista As ListBox, keyini As Currency, keyfin As Currency)
PS_REP02(0) = keyini
PS_REP02(1) = keyfin
llave_rep02.Requery
wlista.Clear
If llave_rep02.EOF Then
 Exit Sub
End If
wlista.Visible = False
Do Until llave_rep02.EOF
 wlista.AddItem llave_rep02!tra_descripcion & String(60, " ") & llave_rep02!TRA_KEY
 If llave_rep02!TRA_FLAG_EXT = "A" Then
   wlista.Selected(wlista.ListCount - 1) = True
 End If
 llave_rep02.MoveNext
Loop
wlista.Visible = True
If wlista.Visible Then wlista.SetFocus

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmensaje.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
If loc_flag_cambio = "A" Then
MsgBox "Debe ´Reiniciar el Sistema´ , para obtener los cambios.", 48, Pub_Titulo
End If
End Sub

Private Sub fraconf_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmensaje.Caption = ""
End Sub

Private Sub lblflag_Click()
ws_nombre = InputBox("Ingresar descripción para etiqueta: ", "Etiqueta", lblflag.Caption)
lblflag.Caption = ws_nombre
End Sub

Private Sub LBLFLAG2_DblClick()
ws_nombre2 = InputBox("Ingresar descripción para etiqueta: ", "Etiqueta", LBLFLAG2.Caption)
LBLFLAG2.Caption = ws_nombre2
End Sub

Private Sub listarep1_GotFocus()
If fraingreso.Visible Then fraingreso.Visible = False
End Sub

Private Sub listarep1_ItemCheck(Item As Integer)
 If WS_FLAG = "A" Then
  Exit Sub
 End If
 ACTUALIZA listarep1, Val(Right(listarep1.text, 6))
End Sub

Public Sub ACTUALIZA(wlista As ListBox, WTRA_KEY As Currency)
 If WTRA_KEY <= 0 Then
   Exit Sub
 End If
 PS_REP01(0) = WTRA_KEY
 llave_rep01.Requery
 If llave_rep01.EOF Then
   MsgBox "Verificar Key de Transacción ..", 48, Pub_Titulo
  Exit Sub
 End If
 Screen.MousePointer = 11
 llave_rep01.Edit
 If wlista.Selected(wlista.ListIndex) Then
   llave_rep01!TRA_FLAG_EXT = "A"
 Else
   llave_rep01!TRA_FLAG_EXT = " "
 End If
 llave_rep01.Update
 Screen.MousePointer = 0

End Sub

Private Sub listarep1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And Trim(listarep1.text) <> "" Then
  pub_mensaje = "Eliminar este Reporte... ¿Desea Continuar... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then
    Exit Sub
  End If
  PUB_CODTRA = Val(Right(listarep1.text, 6))
  pub_cadena = "DELETE TRANSACCION WHERE TRA_KEY = " & PUB_CODTRA
  CN.Execute pub_cadena, rdExecDirect
  listarep1.RemoveItem listarep1.ListIndex
 Exit Sub
End If
If KeyCode = 113 Then
 If Trim(listarep1.text) = "" Then
  Exit Sub
 End If
 SQ_OPER = 1
 WS_FLAG = "E"
 PUB_CODTRA = Val(Right(listarep1.text, 6))
 LIMPIA_DATOS
 LLENA_REPORTE Val(Right(listarep1.text, 6))
 fraingreso.Visible = True
 txtnombre.SetFocus
 Exit Sub
End If
If KeyCode = 115 Then
If Trim(listarep1.text) = "" Then
  Exit Sub
 End If
 Dim wvalor
 wvalor = InputBox("Mover el reporte " & Left(listarep1.text, 25) & "                     en  : (valor entre 1 y 8 )", Pub_Titulo)
 If wvalor = "" Then
  Exit Sub
 End If
 If Val(wvalor) < 1 And Val(wvalor) > 8 Then
   Screen.MousePointer = 0
   MsgBox " Valor incorrecto... ", 48, Pub_Titulo
   Exit Sub
 End If
 If Val(wvalor) = Val(listarep1.Tag) Then
   Screen.MousePointer = 0
   MsgBox " No se puede mover a si mismo... ", 48, Pub_Titulo
   Exit Sub
 End If
 Screen.MousePointer = 11
 REP_CODTRA = Val(Right(listarep1.text, 6))
 If Val(wvalor) = 1 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI1, wsFIN1
   LLENA_LISTA listarep1, wsINI1, wsFIN1
 End If
 If Val(wvalor) = 2 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI2, wsFIN2
   LLENA_LISTA listarep2, wsINI2, wsFIN2
 End If
 If Val(wvalor) = 3 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI3, wsFIN3
   LLENA_LISTA listarep3, wsINI3, wsFIN3
 End If
 If Val(wvalor) = 4 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI4, wsFIN4
   LLENA_LISTA listarep4, wsINI4, wsFIN4
 End If
 If Val(wvalor) = 5 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI5, wsFIN5
   LLENA_LISTA listarep5, wsINI5, wsFIN5
 End If
 If Val(wvalor) = 6 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI6, wsFIN6
   LLENA_LISTA listarep6, wsINI6, wsFIN6
 End If
 If Val(wvalor) = 7 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI7, wsFIN7
   LLENA_LISTA listarep7, wsINI7, wsFIN7
 End If
 If Val(wvalor) = 8 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI8, wsFIN8
   LLENA_LISTA listarep8, wsINI8, wsFIN8
 End If
 LLENA_LISTA listarep1, wsINI1, wsFIN1
End If

End Sub

Private Sub listarep1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmensaje.Caption = "F2 : Modificar Descripción de Reporte / F4 : Mover reporte ...   "
End Sub

Private Sub listarep2_GotFocus()
If fraingreso.Visible Then fraingreso.Visible = False
End Sub

Private Sub listarep2_ItemCheck(Item As Integer)
 If WS_FLAG = "A" Then
  Exit Sub
 End If
 ACTUALIZA listarep2, Val(Right(listarep2.text, 6))

End Sub

Private Sub listarep2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And Trim(listarep2.text) <> "" Then
  pub_mensaje = "Eliminar este Reporte... ¿Desea Continuar... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then
    Exit Sub
  End If
  PUB_CODTRA = Val(Right(listarep2.text, 6))
  pub_cadena = "DELETE TRANSACCION WHERE TRA_KEY = " & PUB_CODTRA
  CN.Execute pub_cadena, rdExecDirect
  listarep2.RemoveItem listarep2.ListIndex
 Exit Sub
End If

If KeyCode = 113 Then
 If Trim(listarep2.text) = "" Then
  Exit Sub
 End If
 SQ_OPER = 2
 WS_FLAG = "E"
 PUB_CODTRA = Val(Right(listarep2.text, 6))
 LIMPIA_DATOS
 LLENA_REPORTE Val(Right(listarep2.text, 6))
 fraingreso.Visible = True
 txtnombre.SetFocus
 Exit Sub
End If

If KeyCode = 115 Then
If Trim(listarep2.text) = "" Then
  Exit Sub
 End If
 Dim wvalor
 wvalor = InputBox("Mover el reporte " & Left(listarep2.text, 25) & "                     en  : (valor entre 1 y 8 )", Pub_Titulo)
 If wvalor = "" Then
  Exit Sub
 End If
 If Val(wvalor) < 1 And Val(wvalor) > 8 Then
   Screen.MousePointer = 0
   MsgBox " Valor incorrecto... ", 48, Pub_Titulo
   Exit Sub
 End If
 If Val(wvalor) = Val(listarep2.Tag) Then
   Screen.MousePointer = 0
   MsgBox " No se puede mover a si mismo... ", 48, Pub_Titulo
   Exit Sub
 End If
 Screen.MousePointer = 11
 REP_CODTRA = Val(Right(listarep2.text, 6))
  If Val(wvalor) = 1 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI1, wsFIN1
   LLENA_LISTA listarep1, wsINI1, wsFIN1
 End If
 If Val(wvalor) = 2 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI2, wsFIN2
   LLENA_LISTA listarep2, wsINI2, wsFIN2
 End If
 If Val(wvalor) = 3 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI3, wsFIN3
   LLENA_LISTA listarep3, wsINI3, wsFIN3
 End If
 If Val(wvalor) = 4 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI4, wsFIN4
   LLENA_LISTA listarep4, wsINI4, wsFIN4
 End If
 If Val(wvalor) = 5 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI5, wsFIN5
   LLENA_LISTA listarep5, wsINI5, wsFIN5
 End If
 If Val(wvalor) = 6 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI6, wsFIN6
   LLENA_LISTA listarep6, wsINI6, wsFIN6
 End If
 If Val(wvalor) = 7 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI7, wsFIN7
   LLENA_LISTA listarep7, wsINI7, wsFIN7
 End If
 If Val(wvalor) = 8 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI8, wsFIN8
   LLENA_LISTA listarep8, wsINI8, wsFIN8
 End If

 LLENA_LISTA listarep2, wsINI2, wsFIN2
End If

End Sub

Private Sub listarep2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmensaje.Caption = "F2 : Modificar Descripción de Reporte / F4 : Mover reporte ...   "
End Sub

Private Sub listarep3_GotFocus()
If fraingreso.Visible Then fraingreso.Visible = False
End Sub

Private Sub listarep3_ItemCheck(Item As Integer)
 If WS_FLAG = "A" Then
  Exit Sub
 End If
 ACTUALIZA listarep3, Val(Right(listarep3.text, 6))

End Sub

Private Sub listarep3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And Trim(listarep3.text) <> "" Then
  pub_mensaje = "Eliminar este Reporte... ¿Desea Continuar... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then
    Exit Sub
  End If
  PUB_CODTRA = Val(Right(listarep3.text, 6))
  pub_cadena = "DELETE TRANSACCION WHERE TRA_KEY = " & PUB_CODTRA
  CN.Execute pub_cadena, rdExecDirect
  listarep3.RemoveItem listarep3.ListIndex
 Exit Sub
End If

If KeyCode = 113 Then
 If Trim(listarep3.text) = "" Then
  Exit Sub
 End If
 SQ_OPER = 3
 WS_FLAG = "E"
 PUB_CODTRA = Val(Right(listarep3.text, 6))
 LIMPIA_DATOS
 LLENA_REPORTE Val(Right(listarep3.text, 6))
 fraingreso.Visible = True
 txtnombre.SetFocus
 Exit Sub
End If

If KeyCode = 115 Then
If Trim(listarep3.text) = "" Then
  Exit Sub
 End If
 Dim wvalor
 wvalor = InputBox("Mover el reporte " & Left(listarep3.text, 25) & "                     en  : (valor entre 1 y 8 )", Pub_Titulo)
 If wvalor = "" Then
  Exit Sub
 End If
 If Val(wvalor) < 1 And Val(wvalor) > 8 Then
   Screen.MousePointer = 0
   MsgBox " Valor incorrecto... ", 48, Pub_Titulo
   Exit Sub
 End If
 If Val(wvalor) = Val(listarep3.Tag) Then
   Screen.MousePointer = 0
   MsgBox " No se puede mover a si mismo... ", 48, Pub_Titulo
   Exit Sub
 End If
 Screen.MousePointer = 11
 REP_CODTRA = Val(Right(listarep3.text, 6))
  If Val(wvalor) = 1 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI1, wsFIN1
   LLENA_LISTA listarep1, wsINI1, wsFIN1
 End If
 If Val(wvalor) = 2 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI2, wsFIN2
   LLENA_LISTA listarep2, wsINI2, wsFIN2
 End If
 If Val(wvalor) = 3 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI3, wsFIN3
   LLENA_LISTA listarep3, wsINI3, wsFIN3
 End If
 If Val(wvalor) = 4 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI4, wsFIN4
   LLENA_LISTA listarep4, wsINI4, wsFIN4
 End If
 If Val(wvalor) = 5 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI5, wsFIN5
   LLENA_LISTA listarep5, wsINI5, wsFIN5
 End If
 If Val(wvalor) = 6 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI6, wsFIN6
   LLENA_LISTA listarep6, wsINI6, wsFIN6
 End If
 If Val(wvalor) = 7 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI7, wsFIN7
   LLENA_LISTA listarep7, wsINI7, wsFIN7
 End If
 If Val(wvalor) = 8 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI8, wsFIN8
   LLENA_LISTA listarep8, wsINI8, wsFIN8
 End If
 LLENA_LISTA listarep3, wsINI3, wsFIN3
End If
End Sub

Private Sub listarep3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmensaje.Caption = "F2 : Modificar Descripción de Reporte / F4 : Mover reporte ...   "
End Sub

Private Sub listarep4_ItemCheck(Item As Integer)
 If WS_FLAG = "A" Then
  Exit Sub
 End If
 ACTUALIZA listarep4, Val(Right(listarep4.text, 6))

End Sub

Private Sub listarep4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And Trim(listarep4.text) <> "" Then
  pub_mensaje = "Eliminar este Reporte... ¿Desea Continuar... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then
    Exit Sub
  End If
  PUB_CODTRA = Val(Right(listarep4.text, 6))
  pub_cadena = "DELETE TRANSACCION WHERE TRA_KEY = " & PUB_CODTRA
  CN.Execute pub_cadena, rdExecDirect
  listarep4.RemoveItem listarep4.ListIndex
 Exit Sub
End If

If KeyCode = 113 Then
 If Trim(listarep4.text) = "" Then
  Exit Sub
 End If
 SQ_OPER = 4
 WS_FLAG = "E"
 PUB_CODTRA = Val(Right(listarep4.text, 6))
 LIMPIA_DATOS
 LLENA_REPORTE Val(Right(listarep4.text, 6))
 fraingreso.Visible = True
 txtnombre.SetFocus
 Exit Sub
End If

If KeyCode = 115 Then
If Trim(listarep4.text) = "" Then
  Exit Sub
 End If
 Dim wvalor
 wvalor = InputBox("Mover el reporte " & Left(listarep4.text, 25) & "                     en  : (valor entre 1 y 8 )", Pub_Titulo)
 If wvalor = "" Then
  Exit Sub
 End If
 If Val(wvalor) < 1 And Val(wvalor) > 8 Then
   Screen.MousePointer = 0
   MsgBox " Valor incorrecto... ", 48, Pub_Titulo
   Exit Sub
 End If
 If Val(wvalor) = Val(listarep4.Tag) Then
   Screen.MousePointer = 0
   MsgBox " No se puede mover a si mismo... ", 48, Pub_Titulo
   Exit Sub
 End If
 Screen.MousePointer = 11
 REP_CODTRA = Val(Right(listarep4.text, 6))
 If Val(wvalor) = 1 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI1, wsFIN1
   LLENA_LISTA listarep1, wsINI1, wsFIN1
 End If
 If Val(wvalor) = 2 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI2, wsFIN2
   LLENA_LISTA listarep2, wsINI2, wsFIN2
 End If
 If Val(wvalor) = 3 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI3, wsFIN3
   LLENA_LISTA listarep3, wsINI3, wsFIN3
 End If
 If Val(wvalor) = 4 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI4, wsFIN4
   LLENA_LISTA listarep4, wsINI4, wsFIN4
 End If
 If Val(wvalor) = 5 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI5, wsFIN5
   LLENA_LISTA listarep5, wsINI5, wsFIN5
 End If
 If Val(wvalor) = 6 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI6, wsFIN6
   LLENA_LISTA listarep6, wsINI6, wsFIN6
 End If
 If Val(wvalor) = 7 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI7, wsFIN7
   LLENA_LISTA listarep7, wsINI7, wsFIN7
 End If
 If Val(wvalor) = 8 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI8, wsFIN8
   LLENA_LISTA listarep8, wsINI8, wsFIN8
 End If
 LLENA_LISTA listarep4, wsINI4, wsFIN4
End If
End Sub

Private Sub listarep4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmensaje.Caption = "F2 : Modificar Descripción de Reporte / F4 : Mover reporte ...   "
End Sub

Private Sub listarep5_GotFocus()
If fraingreso.Visible Then fraingreso.Visible = False
End Sub

Private Sub listarep5_ItemCheck(Item As Integer)
 If WS_FLAG = "A" Then
  Exit Sub
 End If
 ACTUALIZA listarep5, Val(Right(listarep5.text, 6))

End Sub

Private Sub listarep5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And Trim(listarep5.text) <> "" Then
  pub_mensaje = "Eliminar este Reporte... ¿Desea Continuar... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then
    Exit Sub
  End If
  PUB_CODTRA = Val(Right(listarep5.text, 6))
  pub_cadena = "DELETE TRANSACCION WHERE TRA_KEY = " & PUB_CODTRA
  CN.Execute pub_cadena, rdExecDirect
  listarep5.RemoveItem listarep5.ListIndex
 Exit Sub
End If

If KeyCode = 113 Then
 If Trim(listarep5.text) = "" Then
  Exit Sub
 End If
 SQ_OPER = 5
 WS_FLAG = "E"
 PUB_CODTRA = Val(Right(listarep5.text, 6))
 LIMPIA_DATOS
 LLENA_REPORTE Val(Right(listarep5.text, 6))
 fraingreso.Visible = True
 txtnombre.SetFocus
 Exit Sub
End If

If KeyCode = 115 Then
If Trim(listarep5.text) = "" Then
  Exit Sub
 End If
 Dim wvalor
 wvalor = InputBox("Mover el reporte " & Left(listarep5.text, 25) & "                     en  : (valor entre 1 y 8 )", Pub_Titulo)
 If wvalor = "" Then
  Exit Sub
 End If
 If Val(wvalor) < 1 And Val(wvalor) > 8 Then
   Screen.MousePointer = 0
   MsgBox " Valor incorrecto... ", 48, Pub_Titulo
   Exit Sub
 End If
 If Val(wvalor) = Val(listarep5.Tag) Then
   Screen.MousePointer = 0
   MsgBox " No se puede mover a si mismo... ", 48, Pub_Titulo
   Exit Sub
 End If
 Screen.MousePointer = 11
 REP_CODTRA = Val(Right(listarep5.text, 6))
 If Val(wvalor) = 1 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI1, wsFIN1
   LLENA_LISTA listarep1, wsINI1, wsFIN1
 End If
 If Val(wvalor) = 2 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI2, wsFIN2
   LLENA_LISTA listarep2, wsINI2, wsFIN2
 End If
 If Val(wvalor) = 3 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI3, wsFIN3
   LLENA_LISTA listarep3, wsINI3, wsFIN3
 End If
 If Val(wvalor) = 4 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI4, wsFIN4
   LLENA_LISTA listarep4, wsINI4, wsFIN4
 End If
 If Val(wvalor) = 5 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI5, wsFIN5
   LLENA_LISTA listarep5, wsINI5, wsFIN5
 End If
 If Val(wvalor) = 6 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI6, wsFIN6
   LLENA_LISTA listarep6, wsINI6, wsFIN6
 End If
 If Val(wvalor) = 7 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI7, wsFIN7
   LLENA_LISTA listarep7, wsINI7, wsFIN7
 End If
 If Val(wvalor) = 8 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI8, wsFIN8
   LLENA_LISTA listarep8, wsINI8, wsFIN8
 End If
 LLENA_LISTA listarep5, wsINI5, wsFIN5
End If

End Sub

Private Sub listarep5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmensaje.Caption = "F2 : Modificar Descripción de Reporte / F4 : Mover reporte ...   "
End Sub

Private Sub listarep6_GotFocus()
If fraingreso.Visible Then fraingreso.Visible = False
End Sub

Private Sub listarep6_ItemCheck(Item As Integer)
 If WS_FLAG = "A" Then
  Exit Sub
 End If
 ACTUALIZA listarep6, Val(Right(listarep6.text, 6))

End Sub


Private Sub listarep6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And Trim(listarep6.text) <> "" Then
  pub_mensaje = "Eliminar este Reporte... ¿Desea Continuar... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then
    Exit Sub
  End If
  PUB_CODTRA = Val(Right(listarep6.text, 6))
  pub_cadena = "DELETE TRANSACCION WHERE TRA_KEY = " & PUB_CODTRA
  CN.Execute pub_cadena, rdExecDirect
  listarep6.RemoveItem listarep6.ListIndex
 Exit Sub
End If

If KeyCode = 113 Then
 If Trim(listarep6.text) = "" Then
  Exit Sub
 End If
 SQ_OPER = 6
 WS_FLAG = "E"
 PUB_CODTRA = Val(Right(listarep6.text, 6))
 LIMPIA_DATOS
 LLENA_REPORTE Val(Right(listarep6.text, 6))
 fraingreso.Visible = True
 txtnombre.SetFocus
 Exit Sub
End If

If KeyCode = 115 Then
If Trim(listarep6.text) = "" Then
  Exit Sub
 End If
 Dim wvalor
 wvalor = InputBox("Mover el reporte " & Left(listarep6.text, 25) & "                     en  : (valor entre 1 y 8 )", Pub_Titulo)
 If wvalor = "" Then
  Exit Sub
 End If
 If Val(wvalor) < 1 And Val(wvalor) > 8 Then
   Screen.MousePointer = 0
   MsgBox " Valor incorrecto... ", 48, Pub_Titulo
   Exit Sub
 End If
 If Val(wvalor) = Val(listarep6.Tag) Then
   Screen.MousePointer = 0
   MsgBox " No se puede mover a si mismo... ", 48, Pub_Titulo
   Exit Sub
 End If
 Screen.MousePointer = 11
 REP_CODTRA = Val(Right(listarep6.text, 6))
  If Val(wvalor) = 1 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI1, wsFIN1
   LLENA_LISTA listarep1, wsINI1, wsFIN1
 End If
 If Val(wvalor) = 2 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI2, wsFIN2
   LLENA_LISTA listarep2, wsINI2, wsFIN2
 End If
 If Val(wvalor) = 3 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI3, wsFIN3
   LLENA_LISTA listarep3, wsINI3, wsFIN3
 End If
 If Val(wvalor) = 4 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI4, wsFIN4
   LLENA_LISTA listarep4, wsINI4, wsFIN4
 End If
 If Val(wvalor) = 5 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI5, wsFIN5
   LLENA_LISTA listarep5, wsINI5, wsFIN5
 End If
 If Val(wvalor) = 6 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI6, wsFIN6
   LLENA_LISTA listarep6, wsINI6, wsFIN6
 End If
 If Val(wvalor) = 7 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI7, wsFIN7
   LLENA_LISTA listarep7, wsINI7, wsFIN7
 End If
 If Val(wvalor) = 8 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI8, wsFIN8
   LLENA_LISTA listarep8, wsINI8, wsFIN8
 End If
 LLENA_LISTA listarep6, wsINI6, wsFIN6
End If
End Sub

Private Sub listarep6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmensaje.Caption = "F2 : Modificar Descripción de Reporte / F4 : Mover reporte ...   "
End Sub

Private Sub listarep7_GotFocus()
If fraingreso.Visible Then fraingreso.Visible = False
End Sub

Private Sub listarep7_ItemCheck(Item As Integer)
 If WS_FLAG = "A" Then
  Exit Sub
 End If
 ACTUALIZA listarep7, Val(Right(listarep7.text, 6))

End Sub

Private Sub listarep7_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And Trim(listarep7.text) <> "" Then
  pub_mensaje = "Eliminar este Reporte... ¿Desea Continuar... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then
    Exit Sub
  End If
  PUB_CODTRA = Val(Right(listarep7.text, 6))
  pub_cadena = "DELETE TRANSACCION WHERE TRA_KEY = " & PUB_CODTRA
  CN.Execute pub_cadena, rdExecDirect
  listarep7.RemoveItem listarep7.ListIndex
 Exit Sub
End If

If KeyCode = 113 Then
 If Trim(listarep7.text) = "" Then
  Exit Sub
 End If
 SQ_OPER = 7
 WS_FLAG = "E"
 PUB_CODTRA = Val(Right(listarep7.text, 6))
 LIMPIA_DATOS
 LLENA_REPORTE Val(Right(listarep7.text, 6))
 fraingreso.Visible = True
 txtnombre.SetFocus
 Exit Sub
End If

If KeyCode = 115 Then
If Trim(listarep7.text) = "" Then
  Exit Sub
 End If
 Dim wvalor
 wvalor = InputBox("Mover el reporte " & Left(listarep7.text, 25) & "                     en  : (valor entre 1 y 8 )", Pub_Titulo)
 If wvalor = "" Then
  Exit Sub
 End If
 If Val(wvalor) < 1 And Val(wvalor) > 8 Then
   Screen.MousePointer = 0
   MsgBox " Valor incorrecto... ", 48, Pub_Titulo
   Exit Sub
 End If
 If Val(wvalor) = Val(listarep7.Tag) Then
   Screen.MousePointer = 0
   MsgBox " No se puede mover a si mismo... ", 48, Pub_Titulo
   Exit Sub
 End If
 Screen.MousePointer = 11
 REP_CODTRA = Val(Right(listarep7.text, 6))
 If Val(wvalor) = 1 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI1, wsFIN1
   LLENA_LISTA listarep1, wsINI1, wsFIN1
 End If
 If Val(wvalor) = 2 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI2, wsFIN2
   LLENA_LISTA listarep2, wsINI2, wsFIN2
 End If
 If Val(wvalor) = 3 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI3, wsFIN3
   LLENA_LISTA listarep3, wsINI3, wsFIN3
 End If
 If Val(wvalor) = 4 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI4, wsFIN4
   LLENA_LISTA listarep4, wsINI4, wsFIN4
 End If
 If Val(wvalor) = 5 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI5, wsFIN5
   LLENA_LISTA listarep5, wsINI5, wsFIN5
 End If
 If Val(wvalor) = 6 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI6, wsFIN6
   LLENA_LISTA listarep6, wsINI6, wsFIN6
 End If
 If Val(wvalor) = 7 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI7, wsFIN7
   LLENA_LISTA listarep7, wsINI7, wsFIN7
 End If
 If Val(wvalor) = 8 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI8, wsFIN8
   LLENA_LISTA listarep8, wsINI8, wsFIN8
 End If
 LLENA_LISTA listarep7, wsINI7, wsFIN7

End If


End Sub

Private Sub listarep7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmensaje.Caption = "F2 : Modificar Descripción de Reporte / F4 : Mover reporte ...   "
End Sub

Private Sub listarep8_GotFocus()
If fraingreso.Visible Then fraingreso.Visible = False
End Sub

Private Sub listarep8_ItemCheck(Item As Integer)
 If WS_FLAG = "A" Then
  Exit Sub
 End If
 ACTUALIZA listarep8, Val(Right(listarep8.text, 6))

End Sub

Private Sub listarep8_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And Trim(listarep8.text) <> "" Then
  pub_mensaje = "Eliminar este Reporte... ¿Desea Continuar... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then
    Exit Sub
  End If
  PUB_CODTRA = Val(Right(listarep8.text, 6))
  pub_cadena = "DELETE TRANSACCION WHERE TRA_KEY = " & PUB_CODTRA
  CN.Execute pub_cadena, rdExecDirect
  listarep8.RemoveItem listarep8.ListIndex
 Exit Sub
End If

If KeyCode = 113 Then
 If Trim(listarep8.text) = "" Then
  Exit Sub
 End If
 SQ_OPER = 8
 WS_FLAG = "E"
 PUB_CODTRA = Val(Right(listarep8.text, 6))
 LIMPIA_DATOS
 LLENA_REPORTE Val(Right(listarep8.text, 6))
 fraingreso.Visible = True
 txtnombre.SetFocus
 Exit Sub
End If

If KeyCode = 115 Then
 If Trim(listarep8.text) = "" Then
  Exit Sub
 End If
 Dim wvalor
 wvalor = InputBox("Mover el Reporte " & Left(listarep8.text, 25) & "                     en  : (valor entre 1 y 8 )", Pub_Titulo)
 If wvalor = "" Then
  Exit Sub
 End If
 If Val(wvalor) < 1 And Val(wvalor) > 8 Then
   Screen.MousePointer = 0
   MsgBox " Valor incorrecto... ", 48, Pub_Titulo
   Exit Sub
 End If
 If Val(wvalor) = Val(listarep8.Tag) Then
   Screen.MousePointer = 0
   MsgBox " No se puede mover a si mismo... ", 48, Pub_Titulo
   Exit Sub
 End If
 Screen.MousePointer = 11
 REP_CODTRA = Val(Right(listarep8.text, 6))
 If Val(wvalor) = 1 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI1, wsFIN1
   LLENA_LISTA listarep1, wsINI1, wsFIN1
 End If
 If Val(wvalor) = 2 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI2, wsFIN2
   LLENA_LISTA listarep2, wsINI2, wsFIN2
 End If
 If Val(wvalor) = 3 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI3, wsFIN3
   LLENA_LISTA listarep3, wsINI3, wsFIN3
 End If
 If Val(wvalor) = 4 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI4, wsFIN4
   LLENA_LISTA listarep4, wsINI4, wsFIN4
 End If
 If Val(wvalor) = 5 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI5, wsFIN5
   LLENA_LISTA listarep5, wsINI5, wsFIN5
 End If
 If Val(wvalor) = 6 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI6, wsFIN6
   LLENA_LISTA listarep6, wsINI6, wsFIN6
 End If
 If Val(wvalor) = 7 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI7, wsFIN7
   LLENA_LISTA listarep7, wsINI7, wsFIN7
 End If
 If Val(wvalor) = 8 Then
   MOVER_TRA_KEY REP_CODTRA, wsINI8, wsFIN8
   LLENA_LISTA listarep8, wsINI8, wsFIN8
 End If
 LLENA_LISTA listarep8, wsINI8, wsFIN8
End If


End Sub

Private Sub listarep8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmensaje.Caption = "F2 : Modificar Descripción de Reporte / F4 : Mover reporte ...   "
End Sub

Private Sub oparti_DblClick(Index As Integer)
oparti(Index).Value = Not oparti(Index).Value
End Sub

Private Sub opcp_DblClick(Index As Integer)
opcp(Index).Value = Not opcp(Index).Value
End Sub

Private Sub txtarchivo_KeyPress(KeyAscii As Integer)
Dim car
car = UCase(Chr(KeyAscii))
KeyAscii = Asc(car)
If KeyAscii = 13 Then Azul txtformulario, txtformulario

End Sub

Private Sub txtformulario_KeyPress(KeyAscii As Integer)
Dim car
car = UCase(Chr(KeyAscii))
KeyAscii = Asc(car)
End Sub

Private Sub txtnombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Azul txtarchivo, txtarchivo
End Sub

Private Sub txtrep_DblClick(Index As Integer)
'If Trim(LK_CODUSU) <> "SUPERVISOR" Then
' Exit Sub
'End If
If Trim(txtrep(Index).Tag) = "" Then
 Exit Sub
End If
Dim wnombre
wnombre = InputBox("Ingrese la Descripción para  el Titulo :", Pub_Titulo, Trim(txtrep(Index).Caption))
If wnombre = "" Then
  Screen.MousePointer = 0
  Exit Sub
End If
Screen.MousePointer = 11
SQ_OPER = 1
PUB_TIPREG = 99
PUB_NUMTAB = Val(txtrep(Index).Tag)
PUB_CODCIA = "00"
LEER_TAB_LLAVE
If tab_llave.EOF Then
  tab_llave.AddNew
Else
  tab_llave.Edit
End If
  tab_llave!tab_codcia = "00"
  tab_llave!tab_tipreg = 99
  tab_llave!TAB_NUMTAB = Val(txtrep(Index).Tag)
  tab_llave!TAB_NOMLARGO = Left(wnombre, 40)
  tab_llave.Update
  txtrep(Index).Caption = Left(wnombre, 40)
Screen.MousePointer = 0

End Sub
Public Sub MOVER_TRA_KEY(KEY_TRA As Currency, keyini As Currency, keyfin As Currency)
On Error GoTo SALE
Dim WNEWTRA As Currency
Dim wdiref As Integer
wdiref = keyfin - keyini
PS_REP03(0) = keyini
PS_REP03(1) = keyfin
llave_rep03.Requery
If llave_rep03.EOF Then
 'Exit Sub
End If
fila = 0
WNEWTRA = fila + keyini
Do Until llave_rep03.EOF
 fila = fila + 1
 WNEWTRA = fila + keyini
 If llave_rep03!TRA_KEY <> WNEWTRA Then
  GoTo listo
 End If
 llave_rep03.MoveNext
Loop
If fila = wdiref Then
 MsgBox "Tablas de reporte esta llena ", 48, Pub_Titulo
Else
  WNEWTRA = WNEWTRA + 1
  GoTo listo
End If
Screen.MousePointer = 0
Exit Sub

listo:
 PS_REP01(0) = KEY_TRA
 llave_rep01.Requery
 If llave_rep01.EOF Then
   MsgBox "Verificar Key de Transacción ..", 48, Pub_Titulo
  Exit Sub
 End If
 Screen.MousePointer = 11
 llave_rep01.Edit
 llave_rep01!TRA_KEY = WNEWTRA
 llave_rep01.Update
 Screen.MousePointer = 0
 Exit Sub
SALE:
 MsgBox "Verificar Posición", 48, Pub_Titulo
End Sub


Public Sub CAMBIA_NOMBRE(wlista As ListBox, wscodtra As Currency)
Dim wnombre
wnombre = InputBox("Ingrese la Descripción para  el Reporte : (25 Caracteres) ", Pub_Titulo, Left(wlista.text, 25))
If wnombre = "" Then
  Screen.MousePointer = 0
  WS_FLAG = 1
  Exit Sub
End If
PS_REP01(0) = wscodtra
llave_rep01.Requery
If llave_rep01.EOF Then
 MsgBox " Revisar ...codigo no existe en transacciones ..", 48, Pub_Titulo
Else
 llave_rep01.Edit
 llave_rep01!tra_descripcion = Left(wnombre, 25)
 llave_rep01.Update
End If
Screen.MousePointer = 0

End Sub

Private Sub txtrep_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 0 Then
 lblmensaje.Caption = "Doble Click para Modificar titulo ..."
End If
End Sub
Public Function NEW_TRA_KEY(POSI As Integer) As Currency
Dim WNEWTRA As Currency
Dim wdiref As Integer
If POSI = 1 Then
 keyini = wsINI1
 keyfin = wsFIN1
 GoTo dale
ElseIf POSI = 2 Then
 keyini = wsINI2
 keyfin = wsFIN2
 GoTo dale
ElseIf POSI = 3 Then
 keyini = wsINI3
 keyfin = wsFIN3
 GoTo dale
ElseIf POSI = 4 Then
 keyini = wsINI4
 keyfin = wsFIN4
 GoTo dale
ElseIf POSI = 5 Then
 keyini = wsINI5
 keyfin = wsFIN5
 GoTo dale
ElseIf POSI = 6 Then
 keyini = wsINI6
 keyfin = wsFIN6
 GoTo dale
ElseIf POSI = 7 Then
 keyini = wsINI7
 keyfin = wsFIN7
 GoTo dale
ElseIf POSI = 8 Then
 keyini = wsINI8
 keyfin = wsFIN8
 GoTo dale
End If
NEW_TRA_KEY = 0
Exit Function
dale:
PS_REP03(0) = keyini
PS_REP03(1) = keyfin
llave_rep03.Requery
If llave_rep03.EOF Then
 'Exit Sub
End If
wdiref = keyfin - keyini
fila = 0
WNEWTRA = fila + keyini
Do Until llave_rep03.EOF
 fila = fila + 1
 WNEWTRA = fila + keyini
 If llave_rep03!TRA_KEY <> WNEWTRA Then
  GoTo listo
 End If
 llave_rep03.MoveNext
Loop
If fila = wdiref Then
 MsgBox "Tablas de reporte esta llena ", 48, Pub_Titulo
Else
  WNEWTRA = WNEWTRA + 1
  GoTo listo
End If
 Screen.MousePointer = 0

Exit Function

listo:
NEW_TRA_KEY = WNEWTRA
Screen.MousePointer = 0
End Function


Public Sub LIMPIA_DATOS()
oparti(0).Value = False
oparti(1).Value = False
oparti(2).Value = False
chesituacion.Value = 0
Chetipdoc.Value = 0
chefecha.Value = 0
txtnombre.text = ""
txtarchivo.text = ""
txtformulario.text = ""
chefechas.Value = 0
chefechas1.Value = 0
chevendedor.Value = 0
chebancos.Value = 0
opcp(0).Value = False
opcp(1).Value = False
oparti(0).Value = False
oparti(1).Value = False
txtformulario.text = "RCRYSTAL"
checlipro.Value = 0
checolu.Value = 0
chemoneda.Value = 0
cheestado.Value = 0
checar.Value = 0
cheusu.Value = 0

 checia.Value = 0
 chedia.Value = 0
 Chezonas.Value = 0
 chefechas1.Value = 0
 checlipro.Value = 0
 checaa.Value = 0
 chetablas.Value = 0
 cheallog.Value = 0
 checolu.Value = 0
 chefecha.Value = 0
 Chetipdoc.Value = 0
 chesituacion.Value = 0
 chemoneda.Value = 0
 cheestado.Value = 0
 checar.Value = 0
 cheusu.Value = 0
 chesincia.Value = 0
 chenumero.Value = 0
 chemarca.Value = 0
 chestock.Value = 0
 cheped.Value = 0
 cheplancha.Value = 0
 chesubgru.Value = 0
 chelinea.Value = 0
 checal.Value = 0
 checiacar.Value = 0
 chefechapro.Value = 0
 checiaall.Value = 0
 chegiro.Value = 0
 CHEARM.Value = 0
 chetipo.Value = 0
 chearti.Value = 0
End Sub
Public Sub LLENA_REPORTE(WTRA_KEY As Currency)
 If WTRA_KEY <= 0 Then
   Exit Sub
 End If
 PS_REP01(0) = WTRA_KEY
 llave_rep01.Requery
 If llave_rep01.EOF Then
   MsgBox "Verificar Key de Transacción ..", 48, Pub_Titulo
  Exit Sub
 End If
 Screen.MousePointer = 11
 txtnombre.text = Trim(llave_rep01!tra_descripcion)
 txtarchivo.text = Trim(llave_rep01!TRA_L1)
 txtformulario.text = Trim(llave_rep01!TRA_L2)
 cherango.Value = llave_rep01!tra_ACT10
 chefechas.Value = llave_rep01!tra_con1
 chebancos.Value = llave_rep01!TRA_CON2
 chevendedor.Value = llave_rep01!TRA_CON3
 chepto.Value = llave_rep01!tra_rep1
 
 tra_rep1 = chepto.Value
 If llave_rep01!TRA_CON4 = 1 Then
   opcp(0).Value = True
 End If
 If llave_rep01!TRA_CON5 = 1 Then
  opcp(1).Value = True
 End If
 If llave_rep01!tra_con6 = 1 Then
   oparti(0).Value = True
 End If
 If llave_rep01!tra_con7 = 1 Then
   oparti(1).Value = True
 End If
 If llave_rep01!tra_con9 = 1 Then
   oparti(2).Value = True
 End If
 checia.Value = llave_rep01!tra_ACT1
 chedia.Value = llave_rep01!tra_ACT2
 Chezonas.Value = llave_rep01!TRA_CON8
 chefechas1.Value = llave_rep01!tra_con10
 checlipro.Value = llave_rep01!TRA_CON11
 checaa.Value = llave_rep01!tra_con12
 chetablas.Value = llave_rep01!TRA_CON13
 cheallog.Value = llave_rep01!TRA_CON14
 checolu.Value = llave_rep01!TRA_CON15
 chefecha.Value = llave_rep01!TRA_ACT3
 Chetipdoc.Value = llave_rep01!TRA_ACT4
 chesituacion.Value = llave_rep01!TRA_ACT5
 chemoneda.Value = llave_rep01!TRA_ACT6
 cheestado.Value = llave_rep01!TRA_ACT7
 checar.Value = llave_rep01!TRA_ACT8
 cheusu.Value = llave_rep01!TRA_ACT9
 chesincia.Value = Nulo_Valor0(llave_rep01!TRA_ACT11)
 chenumero.Value = Nulo_Valor0(llave_rep01!TRA_ACT12)
 chemarca.Value = Nulo_Valor0(llave_rep01!TRA_ACT13)
 chenat.Value = Nulo_Valor0(llave_rep01!TRA_ACT14)
 chelinea.Value = Nulo_Valor0(llave_rep01!TRA_ACT15)
 checcm.Value = Nulo_Valor0(llave_rep01!TRA_S1)
 chesubgru.Value = Nulo_Valor0(llave_rep01!TRA_S2)
 chestock.Value = Nulo_Valor0(llave_rep01!TRA_S3)
 cheped.Value = Nulo_Valor0(llave_rep01!TRA_S4)
 cheplancha.Value = Nulo_Valor0(llave_rep01!TRA_S5)
 checal.Value = Nulo_Valor0(llave_rep01!TRA_S7)
 checiacar.Value = Nulo_Valor0(llave_rep01!TRA_S6)
 chefechapro.Value = Nulo_Valor0(llave_rep01!TRA_S8)
 checiaall.Value = Nulo_Valor0(llave_rep01!TRA_S9)
 CHEFLAG.Value = Nulo_Valor0(llave_rep01!TRA_S10)
 chegiro.Value = Nulo_Valor0(llave_rep01!TRA_S11)
 CHEARM.Value = Nulo_Valor0(llave_rep01!TRA_S12)
 chetipo.Value = Nulo_Valor0(llave_rep01!TRA_S13)
 chemon.Value = Nulo_Valor0(llave_rep01!TRA_S14)
 CHEMOCAR.Value = Nulo_Valor0(llave_rep01!TRA_s15)
 CHEMOFAR.Value = Nulo_Valor0(llave_rep01!TRA_C1)
 chearti.Value = Nulo_Valor0(llave_rep01!TRA_C2)
 cheprofa.Value = Nulo_Valor0(llave_rep01!TRA_GRU1)
 CHEFLAG2.Value = Nulo_Valor0(llave_rep01!TRA_GRU2)
    
 
  
 lblflag.Caption = Nulo_Valors(llave_rep01!TRA_L3)
 LBLFLAG2.Caption = Nulo_Valors(llave_rep01!TRA_L4)
 
 ws_nombre = Nulo_Valors(llave_rep01!TRA_L3)
 ws_nombre2 = Nulo_Valors(llave_rep01!TRA_L4)
 
 
 
 
 Screen.MousePointer = 0

End Sub

