VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmParGen 
   Caption         =   "Definición de Compañia"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12030
   ControlBox      =   0   'False
   Icon            =   "Frmpargen.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   12030
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
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
      Height          =   625
      Left            =   7080
      Picture         =   "Frmpargen.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   8280
      Width           =   1300
   End
   Begin VB.TextBox txttexto 
      Height          =   285
      Left            =   3720
      MaxLength       =   2
      TabIndex        =   96
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox txtcolor 
      Height          =   285
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   94
      Top             =   720
      Width           =   375
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   0
      TabIndex        =   13
      Top             =   1080
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   10398
      _Version        =   393216
      Tab             =   2
      TabHeight       =   529
      TabCaption(0)   =   "Facturación"
      TabPicture(0)   =   "Frmpargen.frx":058C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "F2"
      Tab(0).Control(1)=   "p_obs"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Archivos/Opcionales "
      TabPicture(1)   =   "Frmpargen.frx":05A8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "F4"
      Tab(1).Control(2)=   "FCAL"
      Tab(1).Control(3)=   "fragrifos"
      Tab(1).Control(4)=   "Frame1"
      Tab(1).Control(5)=   "FRADEVICE"
      Tab(1).Control(6)=   "txtcias"
      Tab(1).Control(7)=   "chepuntos"
      Tab(1).Control(8)=   "cmdComision"
      Tab(1).Control(9)=   "fCOMI"
      Tab(1).Control(10)=   "F3"
      Tab(1).Control(11)=   "Label2(1)"
      Tab(1).Control(12)=   "Label2(0)"
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "Actualizacion de Precios"
      TabPicture(2)   =   "Frmpargen.frx":05C4
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label8"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fracambio"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cboPrecio"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.ComboBox cboPrecio 
         Height          =   315
         ItemData        =   "Frmpargen.frx":05E0
         Left            =   7920
         List            =   "Frmpargen.frx":05F6
         Style           =   2  'Dropdown List
         TabIndex        =   157
         Top             =   960
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         Caption         =   "Archivos de documentos. :"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   128
         Top             =   360
         Width           =   9975
         Begin VB.TextBox t_archi_g 
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
            Left            =   1080
            TabIndex        =   136
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox t_archi_guia 
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
            Left            =   1080
            TabIndex        =   134
            Top             =   1680
            Width           =   1815
         End
         Begin VB.TextBox t_archi_b 
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
            Left            =   1080
            TabIndex        =   133
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox t_archi_f 
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
            Left            =   1080
            TabIndex        =   130
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label ld 
            Alignment       =   1  'Right Justify
            Caption         =   "Guia Int.:"
            Height          =   195
            Index           =   3
            Left            =   -120
            TabIndex        =   135
            Top             =   1200
            Width           =   1140
            WordWrap        =   -1  'True
         End
         Begin VB.Label ld 
            AutoSize        =   -1  'True
            Caption         =   "Guia de Remisión :"
            Height          =   390
            Index           =   2
            Left            =   240
            TabIndex        =   132
            Top             =   1560
            Width           =   735
            WordWrap        =   -1  'True
         End
         Begin VB.Label ld 
            AutoSize        =   -1  'True
            Caption         =   "Boletas :"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   131
            Top             =   720
            Width           =   615
         End
         Begin VB.Label ld 
            AutoSize        =   -1  'True
            Caption         =   "Facturas :"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   129
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Frame F4 
         Caption         =   "Codigo de Articulo"
         Height          =   975
         Left            =   -70440
         TabIndex        =   35
         Top             =   2520
         Visible         =   0   'False
         Width           =   1815
         Begin VB.OptionButton opArti 
            Caption         =   "Alterno / Orig."
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   37
            Top             =   480
            Width           =   1365
         End
         Begin VB.OptionButton opArti 
            Caption         =   "Original"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   1005
         End
      End
      Begin VB.Frame FCAL 
         Caption         =   "Opcion Costo Base :"
         Height          =   855
         Left            =   -70440
         TabIndex        =   30
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         Begin VB.OptionButton OPcosto 
            Caption         =   "Sobre el Valor"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   32
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton OPcosto 
            Caption         =   "Sobre el  (%) "
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   31
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame fragrifos 
         Caption         =   "Sistema :"
         Height          =   975
         Left            =   -72120
         TabIndex        =   58
         Top             =   2520
         Width           =   1575
         Begin VB.CheckBox chegrifos 
            Caption         =   "Funciones de Grifos"
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
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Opciones :"
         Height          =   855
         Left            =   -72120
         TabIndex        =   123
         Top             =   3480
         Width           =   1575
         Begin VB.TextBox txttasa 
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
            Left            =   240
            MaxLength       =   5
            TabIndex        =   124
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Tasa  Venta x Mes  :"
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   125
            Top             =   240
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.TextBox p_obs 
         Height          =   285
         Left            =   -73440
         MaxLength       =   150
         TabIndex        =   126
         Top             =   4860
         Width           =   7815
      End
      Begin VB.Frame fracambio 
         Caption         =   "Actualización de precios de productos en Dolares .:"
         Height          =   3255
         Left            =   120
         TabIndex        =   51
         Top             =   480
         Width           =   7575
         Begin VB.ListBox art_familia 
            Height          =   2310
            Left            =   3720
            Style           =   1  'Checkbox
            TabIndex        =   85
            Top             =   600
            Width           =   3255
         End
         Begin ComctlLib.ProgressBar pb 
            Height          =   255
            Left            =   360
            TabIndex        =   57
            Top             =   2280
            Visible         =   0   'False
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   327682
            Appearance      =   0
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Procesar la Actualización de Precios"
            Height          =   495
            Left            =   480
            TabIndex        =   53
            Top             =   1680
            Width           =   2775
         End
         Begin VB.TextBox txttipo 
            Height          =   375
            Left            =   2280
            TabIndex        =   52
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Seleccione una familia para actualizar o ninguna para todos los productos."
            Height          =   375
            Index           =   1
            Left            =   3720
            TabIndex        =   86
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label5 
            Caption         =   "Label5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   55
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Ultimo Proceso:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   56
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label6 
            Caption         =   "Tipo de Cambio Actual"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   54
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame FRADEVICE 
         Caption         =   "Destino de los Documentos"
         Height          =   1215
         Left            =   -70440
         TabIndex        =   41
         Top             =   4320
         Width           =   5535
         Begin VB.ComboBox cmdimpresoras 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   480
            Width           =   3495
         End
         Begin VB.OptionButton opdevice 
            Caption         =   "Guias"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   44
            Top             =   820
            Width           =   975
         End
         Begin VB.OptionButton opdevice 
            Caption         =   "Boletas"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   43
            Top             =   520
            Width           =   975
         End
         Begin VB.OptionButton opdevice 
            Caption         =   "Facturas "
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   42
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Impresora (segun en panel de control )"
            Height          =   255
            Index           =   0
            Left            =   1440
            TabIndex        =   46
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.TextBox txtcias 
         Height          =   285
         Left            =   -68400
         TabIndex        =   39
         Top             =   3960
         Width           =   1575
      End
      Begin VB.CheckBox chepuntos 
         Caption         =   "Compañia como Punto de Ventas"
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
         Left            =   -68520
         TabIndex        =   38
         Top             =   3480
         Width           =   3495
      End
      Begin VB.CommandButton cmdComision 
         Caption         =   "Insertar / Editar"
         Height          =   495
         Left            =   -66600
         TabIndex        =   33
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Frame fCOMI 
         Caption         =   "Contabilidad"
         Height          =   3015
         Left            =   -74880
         TabIndex        =   25
         Top             =   2520
         Width           =   2655
         Begin VB.TextBox txtparcont 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   91
            Top             =   1680
            Width           =   492
         End
         Begin VB.TextBox par_moneda_con 
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
            Left            =   2160
            MaxLength       =   1
            TabIndex        =   49
            Top             =   1200
            Width           =   375
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Articulo Crear Cta."
            Enabled         =   0   'False
            Height          =   435
            Left            =   240
            TabIndex        =   29
            Top             =   960
            Width           =   1575
         End
         Begin VB.CheckBox Che_contab_p 
            Caption         =   "Proveedor Crear Cta."
            Height          =   435
            Left            =   1440
            TabIndex        =   28
            Top             =   540
            Width           =   1095
         End
         Begin VB.CheckBox Che_contab_c 
            Caption         =   "Clientes Crear Cta."
            Height          =   555
            Left            =   240
            TabIndex        =   27
            Top             =   480
            Width           =   1095
         End
         Begin VB.CheckBox che_contab 
            Caption         =   "Contabilidad"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo de Cia Contab.:"
            Height          =   252
            Index           =   28
            Left            =   120
            TabIndex        =   90
            Top             =   1680
            Width           =   1812
         End
         Begin VB.Label Label1 
            Caption         =   "Moneda para Contabilidad:"
            Height          =   372
            Index           =   25
            Left            =   120
            TabIndex        =   50
            Top             =   1320
            Width           =   2052
         End
      End
      Begin VB.Frame F3 
         Caption         =   "Reportes"
         Height          =   1215
         Left            =   -72120
         TabIndex        =   20
         Top             =   4320
         Visible         =   0   'False
         Width           =   1575
         Begin VB.CheckBox checaja 
            Caption         =   "Caja"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   735
         End
         Begin VB.CheckBox chectaP 
            Caption         =   "Cta. Pagar"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Width           =   1095
         End
         Begin VB.CheckBox chectaC 
            Caption         =   "Cta. Cobrar"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox chebancos 
            Caption         =   "Bancos"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame F2 
         Caption         =   "Definiciones Generales"
         Height          =   5175
         Left            =   -75000
         TabIndex        =   14
         Top             =   420
         Visible         =   0   'False
         Width           =   9855
         Begin VB.CommandButton cmdseries 
            Caption         =   "Ver"
            Height          =   375
            Left            =   4200
            TabIndex        =   141
            Top             =   720
            Width           =   855
         End
         Begin VB.Frame fraseries 
            Caption         =   "fraseries"
            Height          =   3255
            Left            =   4185
            TabIndex        =   140
            Top             =   1155
            Visible         =   0   'False
            Width           =   2895
            Begin VB.ComboBox combotipo 
               Height          =   315
               Left            =   1080
               Style           =   2  'Dropdown List
               TabIndex        =   155
               Top             =   240
               Width           =   1095
            End
            Begin VB.CheckBox cheflag 
               Caption         =   "Activar Serie Predeterminada"
               Height          =   195
               Left            =   120
               TabIndex        =   153
               Top             =   2400
               Width           =   2535
            End
            Begin VB.CheckBox cheini 
               Caption         =   "Activar Numero Inicial"
               Height          =   195
               Left            =   120
               TabIndex        =   152
               Top             =   1080
               Width           =   2055
            End
            Begin VB.TextBox tmax 
               Height          =   285
               Left            =   1080
               MaxLength       =   3
               TabIndex        =   150
               Text            =   "000"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox tfin 
               Height          =   285
               Left            =   1080
               MaxLength       =   9
               TabIndex        =   148
               Text            =   "00000000"
               Top             =   1680
               Width           =   1095
            End
            Begin VB.TextBox tini 
               Height          =   285
               Left            =   1080
               MaxLength       =   9
               TabIndex        =   146
               Text            =   "00000000"
               Top             =   1320
               Width           =   1095
            End
            Begin VB.TextBox tserie 
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
               Left            =   1080
               MaxLength       =   4
               TabIndex        =   145
               Text            =   "0000"
               Top             =   720
               Width           =   615
            End
            Begin VB.CommandButton cancelserie 
               Caption         =   "Cancelar"
               Height          =   435
               Left            =   1680
               TabIndex        =   143
               Top             =   2760
               Width           =   855
            End
            Begin VB.CommandButton grabaserie 
               Caption         =   "Grabar"
               Height          =   435
               Left            =   480
               TabIndex        =   142
               Top             =   2760
               Width           =   855
            End
            Begin VB.Label lser 
               Caption         =   "Tip. Doc.  :"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   154
               Top             =   240
               Width           =   855
            End
            Begin VB.Label lser 
               Caption         =   "Max Item :"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   151
               Top             =   2040
               Width           =   855
            End
            Begin VB.Label lser 
               Caption         =   "Nro. Final :"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   149
               Top             =   1680
               Width           =   855
            End
            Begin VB.Label lser 
               Caption         =   "Nro. Inicial :"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   147
               Top             =   1320
               Width           =   855
            End
            Begin VB.Label lser 
               Caption         =   "Nro. Serie:"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   144
               Top             =   720
               Width           =   855
            End
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Multi Series = "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   2640
            TabIndex        =   139
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtcontado 
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
            Left            =   8640
            MaxLength       =   4
            TabIndex        =   137
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtmorosos 
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
            Left            =   8640
            MaxLength       =   4
            TabIndex        =   121
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox maxguia 
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
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   103
            Top             =   3960
            Width           =   495
         End
         Begin VB.TextBox MaxBoleta 
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
            Left            =   3000
            MaxLength       =   2
            TabIndex        =   102
            Top             =   3960
            Width           =   495
         End
         Begin VB.TextBox Maxfactura 
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
            Left            =   4680
            MaxLength       =   2
            TabIndex        =   101
            Top             =   3960
            Width           =   495
         End
         Begin VB.Frame Frame2 
            Height          =   2535
            Left            =   5640
            TabIndex        =   100
            Top             =   1320
            Width           =   4095
            Begin VB.TextBox serie_ncred 
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
               Left            =   240
               MaxLength       =   4
               TabIndex        =   114
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox ncred_f 
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
               Left            =   240
               MaxLength       =   9
               TabIndex        =   113
               Top             =   1560
               Width           =   975
            End
            Begin VB.TextBox ncred_f_f 
               Height          =   285
               Left            =   240
               MaxLength       =   9
               TabIndex        =   112
               Top             =   2160
               Width           =   975
            End
            Begin VB.CheckBox chencred 
               Caption         =   "Inicializar -Ser. N. Credito"
               Height          =   495
               Left            =   240
               TabIndex        =   111
               Top             =   240
               Width           =   1335
            End
            Begin VB.CheckBox chendeb 
               Caption         =   "Inicializar -Ser. N. Debito"
               Height          =   495
               Left            =   2400
               TabIndex        =   110
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox ndeb_f_f 
               Height          =   285
               Left            =   2400
               MaxLength       =   9
               TabIndex        =   109
               Top             =   2160
               Width           =   975
            End
            Begin VB.TextBox ndeb_f 
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
               Left            =   2400
               MaxLength       =   9
               TabIndex        =   108
               Top             =   1560
               Width           =   975
            End
            Begin VB.TextBox serie_ndeb 
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
               Left            =   2400
               MaxLength       =   4
               TabIndex        =   107
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "N° Nota Credito Inicial"
               Height          =   255
               Index           =   14
               Left            =   120
               TabIndex        =   120
               Top             =   1320
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "N° Nota Credito  Final"
               Height          =   255
               Index           =   15
               Left            =   120
               TabIndex        =   119
               Top             =   1920
               Width           =   1695
            End
            Begin VB.Label Label1 
               Caption         =   "N° Nota Debito  Final"
               Height          =   255
               Index           =   16
               Left            =   2280
               TabIndex        =   118
               Top             =   1920
               Width           =   1695
            End
            Begin VB.Label Label1 
               Caption         =   "N° Nota Debito Inicial"
               Height          =   255
               Index           =   17
               Left            =   2280
               TabIndex        =   117
               Top             =   1320
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "Serie N. Credito."
               Height          =   255
               Index           =   13
               Left            =   240
               TabIndex        =   116
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   "Serie N. Debito."
               Height          =   255
               Index           =   18
               Left            =   2160
               TabIndex        =   115
               Top             =   720
               Width           =   1215
            End
         End
         Begin VB.TextBox txtserie 
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
            Left            =   4080
            MaxLength       =   5
            TabIndex        =   88
            Top             =   1080
            Width           =   615
         End
         Begin VB.CheckBox serie_usu 
            Caption         =   "Facturación x Usuario."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   120
            TabIndex        =   87
            Top             =   720
            Width           =   2415
         End
         Begin VB.TextBox Maxncnd 
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
            Left            =   7680
            MaxLength       =   2
            TabIndex        =   60
            Top             =   3960
            Width           =   375
         End
         Begin VB.TextBox txtmoneda_v 
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
            Left            =   2160
            MaxLength       =   1
            TabIndex        =   47
            Top             =   1080
            Width           =   375
         End
         Begin VB.Frame FDOCU 
            Height          =   2535
            Left            =   120
            TabIndex        =   62
            Top             =   1320
            Width           =   5415
            Begin VB.CheckBox chefactura 
               Caption         =   "Inicializar - Serie Factura"
               Height          =   495
               Left            =   3720
               TabIndex        =   74
               Top             =   240
               Width           =   1335
            End
            Begin VB.CheckBox cheboleta 
               Caption         =   "Inicializar - Serie Boleta "
               Height          =   495
               Left            =   2040
               TabIndex        =   73
               Top             =   240
               Width           =   1215
            End
            Begin VB.CheckBox cheguia 
               Caption         =   "Inicializar - Serie Guia"
               Height          =   495
               Left            =   240
               TabIndex        =   72
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox numfac_f_f 
               Height          =   285
               Left            =   3720
               MaxLength       =   9
               TabIndex        =   71
               Top             =   2160
               Width           =   975
            End
            Begin VB.TextBox numfac_b_f 
               Height          =   285
               Left            =   2040
               MaxLength       =   9
               TabIndex        =   70
               Top             =   2160
               Width           =   975
            End
            Begin VB.TextBox numfac_g_f 
               Height          =   285
               Left            =   240
               MaxLength       =   9
               TabIndex        =   69
               Top             =   2160
               Width           =   975
            End
            Begin VB.TextBox numfac_f 
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
               Left            =   3720
               MaxLength       =   9
               TabIndex        =   68
               Top             =   1560
               Width           =   975
            End
            Begin VB.TextBox serie_f 
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
               Left            =   3720
               MaxLength       =   4
               TabIndex        =   67
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox serie_g 
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
               Left            =   240
               MaxLength       =   4
               TabIndex        =   66
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox numfac_g 
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
               Left            =   240
               MaxLength       =   9
               TabIndex        =   65
               Top             =   1560
               Width           =   975
            End
            Begin VB.TextBox Serie_b 
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
               MaxLength       =   4
               TabIndex        =   64
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox numfac_b 
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
               MaxLength       =   9
               TabIndex        =   63
               Top             =   1560
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "N.Fac Factura Final :"
               Height          =   255
               Index           =   10
               Left            =   3720
               TabIndex        =   83
               Top             =   1920
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "N.Fac Boleta Final "
               Height          =   255
               Index           =   9
               Left            =   2040
               TabIndex        =   82
               Top             =   1920
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "N.Fac Guia Final"
               Height          =   255
               Index           =   8
               Left            =   240
               TabIndex        =   81
               Top             =   1920
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "N.Fac Factura Inicial:"
               Height          =   255
               Index           =   7
               Left            =   3720
               TabIndex        =   80
               Top             =   1320
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "Serie Factura :"
               Height          =   255
               Index           =   3
               Left            =   3720
               TabIndex        =   79
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "N.Fac Guia Inicial "
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   78
               Top             =   1320
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "Serie Guia "
               Height          =   255
               Index           =   5
               Left            =   240
               TabIndex        =   77
               Top             =   720
               Width           =   855
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label1 
               Caption         =   "Serie Boleta "
               Height          =   255
               Index           =   4
               Left            =   2040
               TabIndex        =   76
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "N.Fac Boleta Inicial"
               Height          =   255
               Index           =   2
               Left            =   2040
               TabIndex        =   75
               Top             =   1320
               Width           =   1455
            End
         End
         Begin VB.OptionButton flagfac 
            Caption         =   "Facturación (Talonario) por Compañia"
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
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   3735
         End
         Begin VB.OptionButton flagfac 
            Caption         =   "Facturación (Talonario) por Vendedores"
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
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   3735
         End
         Begin VB.TextBox txtmaxv 
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
            Left            =   8640
            MaxLength       =   3
            TabIndex        =   16
            Top             =   360
            Width           =   615
         End
         Begin VB.CheckBox cheigv 
            Caption         =   "Facturación - I.G.V.  Incluido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5640
            TabIndex        =   15
            Top             =   120
            Width           =   3735
         End
         Begin VB.Label Label1 
            Caption         =   "Dias de gracia para Bloqueo(Contado) :"
            Height          =   255
            Index           =   29
            Left            =   5640
            TabIndex        =   138
            Top             =   1080
            Width           =   2895
         End
         Begin VB.Label Label7 
            Caption         =   "Comentario en Documento :"
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
            TabIndex        =   127
            Top             =   4320
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Dias de gracia para Bloqueo(Credito) :"
            Height          =   255
            Index           =   22
            Left            =   5640
            TabIndex        =   122
            Top             =   720
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "Max. Item G/:"
            Height          =   255
            Index           =   19
            Left            =   240
            TabIndex        =   106
            Top             =   3960
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Max.Item B/:"
            Height          =   255
            Index           =   20
            Left            =   2040
            TabIndex        =   105
            Top             =   3960
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Max.Item F/:"
            Height          =   255
            Index           =   21
            Left            =   3720
            TabIndex        =   104
            Top             =   3960
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Serie de Guias :"
            Height          =   255
            Index           =   27
            Left            =   2880
            TabIndex        =   89
            Top             =   1080
            Width           =   1215
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "La definicón de los documentos está configurada por vendedor."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   960
            Index           =   2
            Left            =   240
            TabIndex        =   84
            Top             =   2040
            Width           =   4800
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            Caption         =   "Max. Item.: N.C / N.D"
            Height          =   255
            Index           =   26
            Left            =   6000
            TabIndex        =   61
            Top             =   3960
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Moneda para Facturacion :"
            Height          =   255
            Index           =   23
            Left            =   240
            TabIndex        =   48
            Top             =   1080
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Maximo de Dias Vencidos :"
            Height          =   255
            Index           =   12
            Left            =   5640
            TabIndex        =   19
            Top             =   360
            Width           =   1935
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de precios Activa:"
         Height          =   195
         Left            =   7920
         TabIndex        =   156
         Top             =   600
         Width           =   1650
      End
      Begin VB.Label Label2 
         Caption         =   "Compañas con la que relaciona."
         Height          =   255
         Index           =   1
         Left            =   -68400
         TabIndex        =   40
         Top             =   3720
         Width           =   2355
      End
      Begin VB.Label Label2 
         Caption         =   "(%) de Comisión y rango de Dias."
         Height          =   435
         Index           =   0
         Left            =   -68520
         TabIndex        =   34
         Top             =   2880
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   645
      Left            =   10440
      TabIndex        =   5
      Top             =   2880
      Width           =   1365
   End
   Begin VB.Timer PARPADEA 
      Interval        =   100
      Left            =   0
      Top             =   6000
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   645
      Left            =   10440
      TabIndex        =   6
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "Ce&rrar"
      Height          =   645
      Left            =   10440
      TabIndex        =   7
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton CmdModificar 
      Caption         =   "&Modificar"
      Height          =   625
      Left            =   10440
      TabIndex        =   3
      Top             =   480
      Width           =   1300
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   645
      Left            =   10440
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Frame F1 
      Caption         =   "Compañia :"
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   9375
      Begin VB.TextBox txtnomcorto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7680
         MaxLength       =   15
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtnombre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         MaxLength       =   40
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox Txt_key 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   2
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Descp. Corto:"
         Height          =   255
         Index           =   11
         Left            =   6600
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción:"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label lblbarraos 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
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
      Left            =   10440
      TabIndex        =   98
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Fondo             Texto:"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   95
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Color de Compañia"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   93
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblcia 
      BackColor       =   &H00FFFFFF&
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
      Height          =   315
      Left            =   4440
      TabIndex        =   92
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label LblMensaje 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   840
      TabIndex        =   11
      Top             =   4920
      Width           =   420
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4914&
      BorderStyle     =   1  'Fixed Single
      Height          =   7095
      Index           =   5
      Left            =   10320
      TabIndex        =   99
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "FrmParGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PASA As Boolean
Dim loc_key As Integer
Dim CU As Integer
Dim parloc_llave As rdoResultset
Dim PSPAR_LOC As rdoQuery
Dim pre_todo As rdoResultset
Dim PSpre_todo As rdoQuery

Dim loc_device As String * 3
Public Function GENERA_PAR() As Integer
Dim valor As Integer
Dim ven_loc As rdoResultset
Dim PSVEN_LOC  As rdoQuery
pub_cadena = "SELECT PAR_CODCIA FROM PARGEN  ORDER BY PAR_CODCIA"
Set PSVEN_LOC = CN.CreateQuery("", pub_cadena)
Set ven_loc = PSVEN_LOC.OpenResultset(rdOpenKeyset, rdConcurValues)
ven_loc.Requery
If ven_loc.EOF Then
 valor = 0
Else
 ven_loc.MoveLast
 valor = ven_loc!PAR_CODCIA
End If
GENERA_PAR = valor + 1

End Function

Public Sub GRABAR_PAR()
Dim walterno As String * 1
Dim wreporte As String * 4
' archivos de documentos

If Left(CmdModificar.Caption, 2) = "&G" Then
   parloc_llave.Edit
   
Else
   parloc_llave.AddNew
End If

parloc_llave!PAR_CODCIA = Format(Val(FrmParGen.Txt_key.text), "00")
parloc_llave!PAR_NOMBRE = FrmParGen.txtnombre.text
parloc_llave!PAR_NOMBRE_CORTO = FrmParGen.txtnomcorto.text
parloc_llave!PAR_G_SERIE = Val(FrmParGen.serie_g.text)
parloc_llave!PAR_G_INICIAL = Val(FrmParGen.numfac_g.text)
parloc_llave!PAR_B_SERIE = Val(FrmParGen.Serie_b.text)
parloc_llave!PAR_B_INICIAL = Val(FrmParGen.numfac_b.text)
parloc_llave!PAR_F_SERIE = Val(FrmParGen.serie_f.text)
parloc_llave!PAR_F_INICIAL = Val(FrmParGen.numfac_f.text)
parloc_llave!PAR_G_FINAL = Val(FrmParGen.numfac_g_f.text)
parloc_llave!PAR_B_FINAL = Val(FrmParGen.numfac_b_f.text)
parloc_llave!PAR_F_FINAL = Val(FrmParGen.numfac_f_f.text)
parloc_llave!PAR_DIAS_VENC = Val(FrmParGen.txtmaxv.text)
parloc_llave!PAR_DIAS_VENC_CC = Val(FrmParGen.txtcontado.text)
parloc_llave!PAR_LISTA = Me.cboPrecio.ListIndex + 1

parloc_llave!PAR_BOL_LINES = Val(FrmParGen.MaxBoleta.text)
parloc_llave!PAR_FAC_LINES = Val(FrmParGen.Maxfactura.text)
parloc_llave!PAR_NOTAS_LINES = Val(FrmParGen.Maxncnd.text)
parloc_llave!PAR_R_SERIE = Val(FrmParGen.txtserie.text)

parloc_llave!PAR_numfac_ncre = Val(FrmParGen.ncred_f.text)
parloc_llave!PAR_SERIE_NCRE = Val(serie_ncred.text)
parloc_llave!PAR_DIAS_LARGE = Val(txtmorosos.text)
parloc_llave!PAR_MONEDA_fac = UCase(txtmoneda_v.text)
parloc_llave!par_moneda_con = UCase(par_moneda_con.text)
parloc_llave!PAR_OBS = Trim(p_obs.text)
parloc_llave!PAR_COLOR = Format(txtcolor.text, "00") & Format(txttexto.text, "00")
parloc_llave!PAR_CIAS = UCase(txtcias.text)
parloc_llave!PAR_TASA_VENTA = Val(txttasa.text)

If chepuntos.Value = 1 Then
  parloc_llave!PAR_PTO_VTA = "A"
Else
  parloc_llave!PAR_PTO_VTA = " "
End If

parloc_llave!PAR_FLAG_NCRE = " "
If chencred.Value = 1 Then
  parloc_llave!PAR_FLAG_NCRE = "A"
End If
parloc_llave!PAR_numfac_nDEB = Val(FrmParGen.ndeb_f.text)
parloc_llave!PAR_SERIE_NDEB = Val(serie_ndeb.text)

parloc_llave!PAR_FLAG_NDEB = " "
If chendeb.Value = 1 Then
  parloc_llave!PAR_FLAG_NDEB = "A"
End If


parloc_llave!PAR_FLAG_G = " "
parloc_llave!PAR_FLAG_B = " "
parloc_llave!PAR_FLAG_F = " "

parloc_llave!PAR_CONTABILIDAD = " "
If che_contab.Value = 1 Then
  parloc_llave!PAR_CONTABILIDAD = "A"
End If
parloc_llave!PAR_CONTA_C = " "
If Che_contab_c.Value = 1 Then
  parloc_llave!PAR_CONTA_C = "A"
End If
parloc_llave!PAR_CONTA_P = " "
If Che_contab_p.Value = 1 Then
  parloc_llave!PAR_CONTA_P = "A"
End If

If cheguia.Value = 1 Then
  parloc_llave!PAR_FLAG_G = "A"
End If
If cheboleta.Value = 1 Then
  parloc_llave!PAR_FLAG_B = "A"
End If
If chefactura.Value = 1 Then
  parloc_llave!PAR_FLAG_F = "A"
End If
parloc_llave!PAR_FLAG_FACTURACION = " "
If flagfac(0).Value Then
 parloc_llave!PAR_FLAG_FACTURACION = "A"
End If
If flagfac(1).Value Then
 parloc_llave!PAR_FLAG_FACTURACION = "V"
End If
If serie_usu.Value = 1 Then
  parloc_llave!PAR_FLAG_FACTURACION = "U"
End If

parloc_llave!PAR_FLAG_ALTERNO = " "
walterno = " "
If opArti(1).Value Then
 parloc_llave!PAR_FLAG_ALTERNO = "A"
 walterno = "A"
End If

If OPcosto(1).Value Then
 parloc_llave!PAR_FLAG_CALCULO = "A"
 LK_FLAG_CALCULO = "A"
Else
 parloc_llave!PAR_FLAG_CALCULO = " "
 LK_FLAG_CALCULO = " "
End If

parloc_llave!PAR_VENTAS_IGV = " "
If cheigv.Value = 1 Then
 parloc_llave!PAR_VENTAS_IGV = "A"
End If

reporte = ""
If chebancos.Value = 1 Then
 reporte = reporte + "A"
Else
 reporte = reporte + " "
End If
If chectaC.Value = 1 Then
 reporte = reporte + "A"
Else
 reporte = reporte + " "
End If
If chectaP.Value = 1 Then
 reporte = reporte + "A"
Else
 reporte = reporte + " "
End If
If checaja.Value = 1 Then
 reporte = reporte + "A"
Else
 reporte = reporte + " "
End If
'parloc_llave!PAR_REPORTES = reporte

parloc_llave!PAR_FLAG_GRIFO = " "
If chegrifos.Value = 1 Then
parloc_llave!PAR_FLAG_GRIFO = "A"
End If
parloc_llave!par_device_fbg = opdevice(0).Tag + opdevice(1).Tag + opdevice(2).Tag
parloc_llave.Update
SQ_OPER = 2
PUB_CODCIA = LK_CODCIA
LEER_PAR_LLAVE
If pac_llave.EOF Then
 pac_llave.AddNew
Else
 pac_llave.Edit
End If
pac_llave!PAC_CODCIA = Format(Val(FrmParGen.Txt_key.text), "00")
pac_llave!PAC_ARCHI_F = Trim(t_archi_f.text)
pac_llave!PAC_ARCHI_B = Trim(t_archi_b.text)
pac_llave!PAC_ARCHI_G = Trim(t_archi_g.text)
pac_llave!PAC_ARCHI_GUIA = Trim(t_archi_guia.text)
pac_llave.Update

If walterno = "A" Then
 LK_FLAG_ORIGINAL = " "
 LK_FLAG_ALTERNO = "A"
 MDIForm1.Toolbar1.Buttons.Item(13).Enabled = True
Else
 LK_FLAG_ORIGINAL = "A"
 LK_FLAG_ALTERNO = " "
 MDIForm1.Toolbar1.Buttons.Item(13).Enabled = False
End If

End Sub
Public Sub MENSAJE_VEN(TEXTO As String)
  LblMensaje.Caption = TEXTO
  PARPADEA.Enabled = True
End Sub

Public Sub LLENA_PAR(ban As Integer)
Dim i As Integer
If ban = 0 Then
       PSPAR_LOC(0) = LK_CODCIA
       parloc_llave.Requery
End If

FrmParGen.Txt_key.text = Trim(Nulo_Valors(parloc_llave!PAR_CODCIA))
FrmParGen.txtnombre.text = Trim(Nulo_Valors(parloc_llave!PAR_NOMBRE))
FrmParGen.txtnomcorto.text = Trim(Nulo_Valors(parloc_llave!PAR_NOMBRE_CORTO))
FrmParGen.serie_g.text = Trim(Nulo_Valors(parloc_llave!PAR_G_SERIE))
FrmParGen.numfac_g.text = Trim(Nulo_Valors(parloc_llave!PAR_G_INICIAL))
FrmParGen.Serie_b.text = Trim(Nulo_Valors(parloc_llave!PAR_B_SERIE))
FrmParGen.numfac_b.text = Trim(Nulo_Valors(parloc_llave!PAR_B_INICIAL))
FrmParGen.serie_f.text = Trim(Nulo_Valors(parloc_llave!PAR_F_SERIE))
FrmParGen.numfac_f.text = Trim(Nulo_Valors(parloc_llave!PAR_F_INICIAL))
FrmParGen.numfac_g_f.text = Trim(Nulo_Valors(parloc_llave!PAR_G_FINAL))
FrmParGen.numfac_b_f.text = Trim(Nulo_Valors(parloc_llave!PAR_B_FINAL))
FrmParGen.numfac_f_f.text = Trim(Nulo_Valors(parloc_llave!PAR_F_FINAL))
serie_ncred.text = Trim(Nulo_Valors(parloc_llave!PAR_SERIE_NCRE))
ncred_f.text = Trim(Nulo_Valors(parloc_llave!PAR_numfac_ncre))
ncred_f_f.text = ""

serie_ndeb.text = Trim(Nulo_Valors(parloc_llave!PAR_SERIE_NDEB))
ndeb_f.text = Trim(Nulo_Valors(parloc_llave!PAR_numfac_nDEB))
ndeb_f_f.text = ""
FrmParGen.MaxBoleta.text = Nulo_Valor0(parloc_llave!PAR_BOL_LINES)
FrmParGen.Maxfactura.text = Nulo_Valor0(parloc_llave!PAR_FAC_LINES)
FrmParGen.Maxncnd.text = Nulo_Valor0(parloc_llave!PAR_NOTAS_LINES)
FrmParGen.txtmorosos.text = Nulo_Valor0(parloc_llave!PAR_DIAS_LARGE)
FrmParGen.txtmoneda_v.text = Nulo_Valor0(parloc_llave!PAR_MONEDA_fac)
FrmParGen.par_moneda_con.text = Nulo_Valor0(parloc_llave!par_moneda_con)
FrmParGen.txtcias.text = Nulo_Valors(parloc_llave!PAR_CIAS)
FrmParGen.txttasa.text = Nulo_Valor0(parloc_llave!PAR_TASA_VENTA)
FrmParGen.txtserie.text = Nulo_Valor0(parloc_llave!PAR_R_SERIE)
p_obs.text = Nulo_Valors(parloc_llave!PAR_OBS)
If UCase(Nulo_Valors(parloc_llave!PAR_PTO_VTA)) = "A" Then
  chepuntos.Value = 1
Else
  chepuntos.Value = 0
End If
Me.cboPrecio.ListIndex = Nulo_Valor0(parloc_llave!PAR_LISTA) - 1 'LISTA DE PRECIOS

FrmParGen.txtcolor.text = Left(Nulo_Valors(parloc_llave!PAR_COLOR), 2)
FrmParGen.txttexto.text = Mid(Nulo_Valors(parloc_llave!PAR_COLOR), 3, 4)

FrmParGen.txtmaxv.text = Nulo_Valor0(parloc_llave!PAR_DIAS_VENC)
FrmParGen.txtcontado.text = Nulo_Valor0(parloc_llave!PAR_DIAS_VENC_CC)
che_contab.Value = 0
If UCase(Nulo_Valors(parloc_llave!PAR_CONTABILIDAD)) = "A" Then
  che_contab.Value = 1
End If
Che_contab_c.Value = 0
If UCase(Nulo_Valors(parloc_llave!PAR_CONTA_C)) = "A" Then
  Che_contab_c.Value = 1
End If
Che_contab_p.Value = 0
If UCase(Nulo_Valors(parloc_llave!PAR_CONTA_P)) = "A" Then
  Che_contab_p.Value = 1
End If
If che_contab.Value = 0 Then
 Che_contab_c.Enabled = False
 Che_contab_p.Enabled = False
End If
chendeb.Value = 0
If UCase(Nulo_Valors(parloc_llave!PAR_FLAG_NDEB)) = "A" Then
  chendeb.Value = 1
End If

chencred.Value = 0
If UCase(Nulo_Valors(parloc_llave!PAR_FLAG_NCRE)) = "A" Then
  chencred.Value = 1
End If
cheguia.Value = 0
cheboleta.Value = 0
chefactura.Value = 0
If UCase(Nulo_Valors(parloc_llave!PAR_FLAG_G)) = "A" Then
  cheguia.Value = 1
End If
If UCase(Nulo_Valors(parloc_llave!PAR_FLAG_B)) = "A" Then
  cheboleta.Value = 1
End If
If UCase(Nulo_Valors(parloc_llave!PAR_FLAG_F)) = "A" Then
  chefactura.Value = 1
End If
opArti(0).Value = False
opArti(1).Value = False
If parloc_llave!PAR_FLAG_ALTERNO <> "A" Then
 opArti(0).Value = True
End If
If parloc_llave!PAR_FLAG_ALTERNO = "A" Then
 opArti(1).Value = True
End If

OPcosto(0).Value = False
OPcosto(1).Value = False
If parloc_llave!PAR_FLAG_CALCULO <> "A" Then
 OPcosto(0).Value = True
End If
If parloc_llave!PAR_FLAG_CALCULO = "A" Then
 OPcosto(1).Value = True
End If

flagfac(0).Value = False
flagfac(1).Value = False
If parloc_llave!PAR_FLAG_FACTURACION = "A" Then
 flagfac(0).Value = True
End If
If parloc_llave!PAR_FLAG_FACTURACION <> "A" Then
 flagfac(1).Value = True
End If
If parloc_llave!PAR_FLAG_FACTURACION = "U" Then
   serie_usu.Value = 1
End If
'If Right(Left(parloc_llave!PAR_REPORTES, 1), 1) = "A" Then
' chebancos.Value = 1
'End If
'If Right(Left(parloc_llave!PAR_REPORTES, 2), 1) = "A" Then
' chectaC.Value = 1
'End If
'If Right(Left(parloc_llave!PAR_REPORTES, 3), 1) = "A" Then
' chectaP.Value = 1
'End If
'If Right(Left(parloc_llave!PAR_REPORTES, 4), 1) = "A" Then
' checaja.Value = 1
'End If
cheigv.Value = 0
If UCase(Nulo_Valors(parloc_llave!PAR_VENTAS_IGV)) = "A" Then
  cheigv.Value = 1
End If

If UCase(Nulo_Valors(parloc_llave!PAR_FLAG_GRIFO)) = "A" Then
  chegrifos.Value = 1
Else
  chegrifos.Value = 0
End If

opdevice(0).Tag = Mid(Trim(Nulo_Valors(parloc_llave!par_device_fbg)), 1, 1)
opdevice(1).Tag = Mid(Trim(Nulo_Valors(parloc_llave!par_device_fbg)), 2, 1)
opdevice(2).Tag = Mid(Trim(Nulo_Valors(parloc_llave!par_device_fbg)), 3, 1)

SQ_OPER = 2
PUB_CODCIA = LK_CODCIA
LEER_PAR_LLAVE
If Not pac_llave.EOF Then
 t_archi_f.text = Trim(pac_llave!PAC_ARCHI_F)
 t_archi_b.text = Trim(pac_llave!PAC_ARCHI_B)
 t_archi_g.text = Trim(pac_llave!PAC_ARCHI_G)
 t_archi_guia.text = Trim(pac_llave!PAC_ARCHI_GUIA)
End If



End Sub
Public Sub LIMPIA_PAR()
FrmParGen.txttasa.text = ""
txtcias.text = ""
txtmoneda_v.text = ""
chepuntos.Value = 0
Txt_key.text = ""
txtnombre.text = ""
txtnomcorto.text = ""
serie_g.text = ""
numfac_g.text = ""
Serie_b.text = ""
numfac_b.text = ""
serie_f.text = ""
numfac_f.text = ""
numfac_g_f.text = ""
numfac_b_f.text = ""
numfac_f_f.text = ""
cheguia.Value = 0
cheboleta.Value = 0
chefactura.Value = 0
flagfac(0).Value = False
flagfac(1).Value = False
cheigv.Value = 0
opArti(0).Value = False
opArti(1).Value = False
OPcosto(0).Value = False
OPcosto(1).Value = False
txtmaxv.text = ""
che_contab.Value = 0
Che_contab_c.Value = 0
Che_contab_p.Value = 0
Che_contab_c.Value = 0
chencred.Value = 0
serie_ncred.text = ""
ncred_f.text = ""
ncred_f_f.text = ""
chendeb.Value = 0
serie_ndeb.text = ""
ndeb_f.text = ""
ndeb_f_f.text = ""
maxguia.text = ""
Maxfactura.text = ""
MaxBoleta.text = ""
Maxncnd.text = ""
par_moneda_con.text = ""
chegrifos.Value = 0
serie_usu.Value = 0
txtcolor.text = ""
txttexto.text = ""
t_archi_f.text = ""
t_archi_b.text = ""
t_archi_g.text = ""
t_archi_guia.text = ""
Me.cboPrecio.ListIndex = -1
End Sub

Private Sub che_contab_Click()
If che_contab.Value = 1 Then
Che_contab_p.Enabled = True
Che_contab_c.Enabled = True
Che_contab_p.Value = 0
Che_contab_c.Value = 0
Else
Che_contab_p.Enabled = False
Che_contab_c.Enabled = False
Che_contab_p.Value = 0
Che_contab_c.Value = 0
End If

End Sub

Private Sub cheboleta_Click()
If Serie_b.Enabled Then
 Serie_b.SetFocus
End If
End Sub

Private Sub chefactura_Click()
If serie_f.Enabled Then
 serie_f.SetFocus
End If
End Sub

Private Sub cheguia_Click()
If serie_g.Enabled Then
 serie_g.SetFocus
End If
End Sub

Private Sub cmdagregar_Click()
'On Error GoTo ESCAPA
If Left(CmdAgregar.Caption, 2) = "&A" Then
    CmdAgregar.Caption = "&Grabar"
    cmdCancelar.Enabled = True
    CmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    LIMPIA_PAR
    DESBLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, chebancos, chectaC, chectaP
    DESBLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, txtnomcorto, flagfac(0), flagfac(1), checaja
    DESBLOQUEA_TEXT opArti(0), opArti(1), OPcosto(0), OPcosto(1), txtmaxv, fCOMI, chencred, serie_ncred, ncred_f, ncred_f_f
    DESBLOQUEA_TEXT serie_ndeb, ndeb_f, ndeb_f_f, chendeb, chendeb, MaxBoleta, Maxfactura, txtmorosos, chepuntos, chepuntos
    DESBLOQUEA_TEXT txtcias, txttasa, cmdimpresoras, par_moneda_con, chegrifos, Maxncnd, serie_usu, txtserie
    DESBLOQUEA_TEXT t_archi_f, t_archi_b, t_archi_g, t_archi_guia, txtcontado
    cheigv.Enabled = True
    FrmParGen.Txt_key = GENERA_PAR
    FrmParGen.Txt_key.Locked = False
    FrmParGen.txtnombre.SetFocus
    'AGREGAMOS EN BLANCO
Else
   If FrmParGen.txtnombre.text = "" Or Len(FrmParGen.txtnombre.text) = 0 Then
       MsgBox "Ingrese Nombre de Vendedor ..!!!", 48, Pub_Titulo
       Azul txtnombre, txtnombre
       Exit Sub
   End If
   If Val(Txt_key.text) < 0 Then
     MsgBox "Codigo de Cia Invalido ...", 48, Pub_Titulo
     Azul Txt_key, Txt_key
     Exit Sub
   End If
   PSPAR_LOC(0) = Format(Txt_key.text, "00")
   parloc_llave.Requery
   If Not parloc_llave.EOF Then
     MsgBox "La Compañia EXISTE en el Sistema Carge El Formulario Nuevamente.", 48, Pub_Titulo
     Unload FrmParGen
     Exit Sub
   End If
   Screen.MousePointer = 11
   GRABAR_PAR
   MENSAJE_VEN "Bancos , AGREGADO... "
   CmdAgregar.Caption = "&Agregar"
   cmdEliminar.Enabled = True
   CmdModificar.Enabled = True
   LIMPIA_PAR
   Screen.MousePointer = 0
   MsgBox "Se Agrego una Nueva Compañia , Debe dar derechos a los usuarios y Reiniciar el Sistema.", 48, Pub_Titulo
   Unload FrmParGen
   Exit Sub
End If
   
End Sub

Private Sub cmdAgregar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    
End If

End Sub

Private Sub cmdcancelar_Click()
If Left(CmdAgregar.Caption, 2) = "&A" And Left(CmdModificar.Caption, 2) = "&M" Then
     Exit Sub
End If
     Screen.MousePointer = 11
     If Left(CmdModificar.Caption, 2) = "&G" Then
        CmdModificar.Caption = "&Modificar"
        LLENA_PAR 1
        BLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, chebancos, chectaC, chectaP
        BLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, txtnomcorto, flagfac(0), flagfac(1), checaja
        BLOQUEA_TEXT opArti(0), opArti(1), OPcosto(0), OPcosto(1), txtmaxv, fCOMI, chencred, serie_ncred, ncred_f, ncred_f_f
        BLOQUEA_TEXT serie_ndeb, ndeb_f, ndeb_f_f, chendeb, MaxBoleta, Maxfactura, txtmorosos, chepuntos
        BLOQUEA_TEXT txtcias, txttasa, cmdimpresoras, par_moneda_con, chegrifos, Maxncnd, serie_usu, txtserie
        BLOQUEA_TEXT t_archi_f, t_archi_b, t_archi_g, t_archi_guia, txtcontado
        cheigv.Enabled = False
        Txt_key.Locked = True
     Else
        CmdAgregar.Caption = "&Agregar"
        LLENA_PAR 1
        BLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, chebancos, chectaC, chectaP
        BLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, txtnomcorto, flagfac(0), flagfac(1), checaja
        BLOQUEA_TEXT opArti(0), opArti(1), OPcosto(0), OPcosto(1), txtmaxv, fCOMI, chencred, serie_ncred, ncred_f, ncred_f_f
        BLOQUEA_TEXT serie_ndeb, ndeb_f, ndeb_f_f, chendeb, MaxBoleta, Maxfactura, txtmorosos, chepuntos
        BLOQUEA_TEXT txtcias, txttasa, cmdimpresoras, par_moneda_con, chegrifos, Maxncnd, serie_usu, txtserie
        BLOQUEA_TEXT t_archi_f, t_archi_b, t_archi_g, t_archi_guia, txtcontado
        cheigv.Enabled = False
        Txt_key.Locked = True
     End If
     CmdCerrar.Caption = "&Cerrar"
     cmdCancelar.Enabled = True
     CmdAgregar.Enabled = True
     CmdModificar.Enabled = True
     cmdEliminar.Enabled = True
     Txt_key.Enabled = True
     MENSAJE_VEN "Proceso Cancelado... !!!    "
     Txt_key.SetFocus
     Screen.MousePointer = 0

End Sub

Private Sub cmdcerrar_Click()
ws_conta = 0
Unload FrmParGen

End Sub

Private Sub cmdCerrar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    FrmParGen.Txt_key.SetFocus
End If

End Sub

Private Sub cmdComision_Click()
Dim RES As String
RES = InputBox("Para mostrar la Tabla de Rango de Minorista digitar 1 , para la de Mayorista 2 ", "Ingrese Datos", "")
If RES = "" Then Exit Sub
If Not (Val(RES) = 1 Or Val(RES) = 2) Then Exit Sub

If Val(RES) = 1 Then
  PUB_TIPREG = 444
Else
  PUB_TIPREG = 445
End If
PUB_CODCIA = LK_CODCIA
Load FrmDatArti
If Val(RES) = 1 Then
  FrmDatArti.Caption = "Rango de Minoristas, Comisiones " & PUB_TIPREG
Else
  FrmDatArti.Caption = "Rango de Mayoristas, Comisiones " & PUB_TIPREG
End If
FrmDatArti.Show 1
DoEvents
cmdComision.SetFocus
'cmdcerrar.SetFocus

End Sub

Private Sub cmdEliminar_Click()
If Len(Txt_key) = 0 Or Len(txtnombre) = 0 Then
   MENSAJE_VEN "NO a seleccionado NADA ... !"
   Exit Sub
End If
  pub_mensaje = " ¿Desea Eliminar el Registro... ?, al Eliminar, debe quitar el Permiso al usuarios que esta usandolo"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbYes Then   ' El usuario eligió
    Screen.MousePointer = 11
    parloc_llave.Delete
    Txt_key.text = ""
    Txt_key.Locked = False
    LIMPIA_PAR
    MENSAJE_VEN "Registro   ELIMINADO ... "
    MsgBox "Compañia Eliminada , se procede a salir de Sistema, vuelva a ingresar y cambie de compañia", 48, ws_titulo
    ' salir
    Screen.MousePointer = 11
    DoEvents
    MDIForm1.WindowState = 1
    MDIForm1.Caption = ""
    EN.Close
    CN.Close
    Screen.MousePointer = 0
    End
   Exit Sub
  End If
  Screen.MousePointer = 0
End Sub


Private Sub cmdimpresoras_Click()
Dim tem As String * 3
If loc_key = 99 Then Exit Sub
tem = loc_device
If opdevice(0).Value Then
  opdevice(0).Tag = Trim(Right(cmdimpresoras.text, 4))
ElseIf opdevice(1).Value Then
  opdevice(1).Tag = Trim(Right(cmdimpresoras.text, 4))
ElseIf opdevice(2).Value Then
  opdevice(2).Tag = Trim(Right(cmdimpresoras.text, 4))
End If

End Sub

Private Sub CmdModificar_Click()
If Len(Txt_key) = 0 Then
   MENSAJE_VEN "NO a seleccionado NADA ... !"
   Exit Sub
End If
If Left(CmdModificar.Caption, 2) = "&M" Then
    CmdModificar.Caption = "&Grabar"
    CmdAgregar.Enabled = False
    cmdEliminar.Enabled = False
    cmdCancelar.Enabled = True
    Txt_key.Locked = True
    DESBLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, chebancos, chectaC, chectaP
    DESBLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, txtnomcorto, flagfac(0), flagfac(1), checaja
    DESBLOQUEA_TEXT opArti(0), opArti(1), OPcosto(0), OPcosto(1), txtmaxv, fCOMI, chencred, serie_ncred, ncred_f, ncred_f_f
    DESBLOQUEA_TEXT serie_ndeb, ndeb_f, ndeb_f_f, chendeb, MaxBoleta, Maxfactura, txtmorosos, chepuntos
    DESBLOQUEA_TEXT txtcias, txttasa, cmdimpresoras, par_moneda_con, chegrifos, Maxncnd, serie_usu, txtserie
    DESBLOQUEA_TEXT t_archi_f, t_archi_b, t_archi_g, t_archi_guia, txtcontado
    cheigv.Enabled = True
    txtnombre.SetFocus
Else
    '*Grabar las modificaciones
    If txtnombre.text = "" Or Len(txtnombre.text) = 0 Then
         MsgBox " Nombre Invalido ....", 48, Pub_Titulo
         Exit Sub
    End If
     Screen.MousePointer = 11
     GRABAR_PAR
     MENSAJE_VEN "Registro , MODIFICADO... "
     CmdModificar.Caption = "&Modificar"
     cmdCancelar.Enabled = True
     CmdAgregar.Enabled = True
     cmdEliminar.Enabled = True
     Txt_key.Locked = True
     BLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, chebancos, chectaC, chectaP
     BLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, txtnomcorto, flagfac(0), flagfac(1), checaja
     BLOQUEA_TEXT opArti(0), opArti(1), OPcosto(0), OPcosto(1), txtmaxv, fCOMI, chencred, serie_ncred, ncred_f, ncred_f_f
     BLOQUEA_TEXT serie_ndeb, ndeb_f, ndeb_f_f, chendeb, MaxBoleta, Maxfactura, txtmorosos, chepuntos
     BLOQUEA_TEXT txtcias, txttasa, cmdimpresoras, par_moneda_con, chegrifos, Maxncnd, serie_usu, txtserie
     BLOQUEA_TEXT t_archi_f, t_archi_b, t_archi_g, t_archi_guia, txtcontado
     cheigv.Enabled = False
     Screen.MousePointer = 0
     MsgBox "Datos de la Compañia ha cambiado, se Reinicia a Cia:  " & (txtnombre.text) & ".", 48, Pub_Titulo
     ACTUALIZA_CIA LK_CODCIA
     Unload FrmParGen
     Exit Sub
End If

End Sub

Private Sub cmdseries_Click()
combotipo.Clear
combotipo.AddItem "G=Guias"
combotipo.AddItem "B=Boletas"
combotipo.AddItem "F=Facturas"
combotipo.AddItem "N=Not.Credito"
combotipo.AddItem "D=Not.Debito"
combotipo.ListIndex = 0
fraseries.Visible = True
combotipo.SetFocus

End Sub

Private Sub combotipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul tserie, tserie
End If
End Sub

Private Sub Command1_Click()
Dim km_flag As String
    If Val(txttipo.text) <= 0 Then
     MsgBox "Verificar el Tipo de Cambio para el proceso.No Procede.", 48, Pub_Titulo
     Exit Sub
    End If
    MODO1 = " AND ("
    km_flag = ""
    art_familia.Visible = False
    For fila = 0 To art_familia.ListCount - 1
      art_familia.ListIndex = fila
      If art_familia.Selected(fila) Then
        MODO1 = MODO1 + "ART_FAMILIA = " & Trim(Right(art_familia.text, 6)) & " OR "
        km_flag = "A"
       End If
    Next fila
    art_familia.Visible = True
    If km_flag = "A" Then
      MODO1 = Mid(MODO1, 1, Len(MODO1) - 3) & " )"
      pub_mensaje = " ¿Desea Actualizar los Precios a las Familias seleccionadas... ?"
    Else
      MODO1 = ""
      pub_mensaje = " ¿Desea Actualizar los Precios a todos los productos...?"
    End If
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
       Exit Sub
    End If
    pub_cadena = "SELECT ART_FAMILIA, PRE_PRE1, PRE_PRE2, PRE_PRE3, PRE_PRE4, PRE_PRE5, PRE_PRE11, PRE_PRE22, PRE_PRE33, PRE_PRE44, PRE_PRE55 FROM ARTI, PRECIOS WHERE PRE_CODCIA = ART_CODCIA AND ART_KEY = PRE_CODART AND ART_KEY <> 0 AND PRE_CODCIA = ? " + MODO1
    Set PSpre_todo = CN.CreateQuery("", pub_cadena)
    PSpre_todo.rdoParameters(0) = 0
    Set pre_todo = PSpre_todo.OpenResultset(rdOpenKeyset, rdConcurValues)
    PSpre_todo.rdoParameters(0) = LK_CODCIA
    pre_todo.Requery
    If pre_todo.EOF Then
      MsgBox "No hay Productos para Actualizar.", 48, Pub_Titulo
      GoTo PASA
    End If
    pb.Visible = True
    DoEvents
    pb.Min = 0
    pb.Max = pre_todo.RowCount
    pb.Value = 0
    Do Until pre_todo.EOF
       pb.Value = pb.Value + 1
       pre_todo.Edit
       pre_todo!pre_pre1 = redondea(pre_todo!pre_pre11 * Val(txttipo.text))
       pre_todo!pre_pre2 = redondea(pre_todo!pre_pre22 * Val(txttipo.text))
       pre_todo!pre_pre3 = redondea(pre_todo!pre_pre33 * Val(txttipo.text))
       pre_todo!pre_pre4 = redondea(pre_todo!pre_pre44 * Val(txttipo.text))
       pre_todo!pre_pre5 = redondea(pre_todo!pre_pre55 * Val(txttipo.text))
       pre_todo.Update
       pre_todo.MoveNext
    Loop
        
    parloc_llave.Requery
    parloc_llave.Edit
    parloc_llave!PAR_FECHA_PRECIOS = LK_FECHA_DIA
    parloc_llave!PAR_TASA_VENTA = txttipo.text
    parloc_llave.Update
    pb.Visible = False
    MsgBox "Proceso Terminado...", 48, Pub_Titulo

    Exit Sub
PASA:

End Sub


Private Sub flagfac_Click(Index As Integer)
If Index = 0 Then
 FDOCU.Visible = True
Else
 FDOCU.Visible = False
End If
End Sub

Private Sub Form_Activate()
F1.Visible = True
F2.Visible = True
F4.Visible = True
FCAL.Visible = True
End Sub

Private Sub Form_Load()
pub_cadena = "SELECT * FROM PARGEN WHERE PAR_CODCIA = ? ORDER BY  PAR_NOMBRE"
Set PSPAR_LOC = CN.CreateQuery("", pub_cadena)
PSPAR_LOC(0) = 0
Set parloc_llave = PSPAR_LOC.OpenResultset(rdOpenKeyset, rdConcurValues)

loc_key = 0
LIMPIA_PAR
Txt_key.text = LK_CODCIA
PSPAR_LOC(0) = LK_CODCIA
parloc_llave.Requery
If parloc_llave.EOF Then
  Azul Txt_key, Txt_key
  MsgBox "Compañia NO EXSITE en PARGEN, consulte al Administrador, Procede a Salir de Sistema", 48, Pub_Titulo
  Screen.MousePointer = 11
  DoEvents
  MDIForm1.WindowState = 1
  MDIForm1.Caption = ""
  EN.Close
  CN.Close
  Screen.MousePointer = 0
  End
  Exit Sub
End If
cmdCancelar.Enabled = True
cheigv.Enabled = False
BLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, chebancos, chectaC, chectaP
BLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, txtnomcorto, flagfac(0), flagfac(1), checaja
BLOQUEA_TEXT opArti(0), opArti(1), OPcosto(0), OPcosto(1), txtmaxv, fCOMI, chencred, serie_ncred, ncred_f, ncred_f_f
BLOQUEA_TEXT serie_ndeb, ndeb_f, ndeb_f_f, chendeb, MaxBoleta, Maxfactura, txtmorosos, txtmorosos, chepuntos
BLOQUEA_TEXT txtcias, txttasa, cmdimpresoras, par_moneda_con, chegrifos, Maxncnd, serie_usu, txtserie
BLOQUEA_TEXT t_archi_f, t_archi_b, t_archi_g, t_archi_guia, txtcontado

LLENA_PAR 0
Txt_key.Enabled = True
Txt_key.Locked = True

Dim P As Printer
fila = 0
For Each P In Printers
    cmdimpresoras.AddItem P.DeviceName & String(80, " ") & fila
    fila = fila + 1
Next P
fracambio.Visible = False
'If Trim(LK_EMP) = "3AA" Then
Label5.Caption = Format(parloc_llave!PAR_FECHA_PRECIOS, "dd/mm/yyyy")
txttipo.text = Format(parloc_llave!PAR_TASA_VENTA, "0.0000")
fracambio.Visible = True
LLENADO_FAM
'End If

If LK_FLAG_FACTURACION = "A" Then
  FDOCU.Visible = True
Else
  FDOCU.Visible = False
End If

Screen.MousePointer = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
ws_conta = 0
End Sub

Public Sub BLOQUEA_TEXT(Optional o1, Optional o2, Optional o3, Optional o4, Optional o5, Optional o6, Optional o7, Optional o8, Optional o9, Optional o10)
'** BLOQUEA TEXTBOX  CANTIDAD DE OBJECTOS **
If Not IsMissing(o1) Then
 o1.Enabled = False
' o1.BackColor = QBColor(7)
End If
If Not IsMissing(o2) Then
 o2.Enabled = False
 'o2.BackColor = QBColor(7)
End If
If Not IsMissing(o3) Then
 o3.Enabled = False
 'o3.BackColor = QBColor(7)
End If
If Not IsMissing(o4) Then
 o4.Enabled = False
 'o4.BackColor = QBColor(7)
End If
If Not IsMissing(o5) Then
 o5.Enabled = False
 'o5.BackColor = QBColor(7)
End If
If Not IsMissing(o6) Then
 o6.Enabled = False
 'o6.BackColor = QBColor(7)
End If
If Not IsMissing(o7) Then
 o7.Enabled = False
 'o7.BackColor = QBColor(7)
End If
If Not IsMissing(o8) Then
 o8.Enabled = False
 'o8.BackColor = QBColor(7)
End If
If Not IsMissing(o9) Then
 o9.Enabled = False
 'o9.BackColor = QBColor(7)
End If
If Not IsMissing(o10) Then
 o10.Enabled = False
 'o10.BackColor = QBColor(7)
End If
End Sub
Public Sub DESBLOQUEA_TEXT(Optional o1, Optional o2, Optional o3, Optional o4, Optional o5, Optional o6, Optional o7, Optional o8, Optional o9, Optional o10)
'** BLOQUEA TEXTBOX  CANTIDAD DE OBJECTOS **
If Not IsMissing(o1) Then
 o1.Enabled = True
' o1.BackColor = QBColor(15)
End If
If Not IsMissing(o2) Then
 o2.Enabled = True
' o2.BackColor = QBColor(15)
End If
If Not IsMissing(o3) Then
 o3.Enabled = True
' o3.BackColor = QBColor(15)
End If
If Not IsMissing(o4) Then
 o4.Enabled = True
' o4.BackColor = QBColor(15)
End If
If Not IsMissing(o5) Then
 o5.Enabled = True
' o5.BackColor = QBColor(15)
End If
If Not IsMissing(o6) Then
 o6.Enabled = True
' o6.BackColor = QBColor(15)
End If
If Not IsMissing(o7) Then
 o7.Enabled = True
' o7.BackColor = QBColor(15)
End If
If Not IsMissing(o8) Then
 o8.Enabled = True
' o8.BackColor = QBColor(15)
End If
If Not IsMissing(o9) Then
 o9.Enabled = True
' o9.BackColor = QBColor(15)
End If
If Not IsMissing(o10) Then
 o10.Enabled = True
' o10.BackColor = QBColor(15)
End If
End Sub

Private Sub grabaserie_Click()
 If ser_llave.EOF Then
   ser_llave.AddNew
 Else
   ser_llave.Edit
 End If
 ser_llave!ser_codcia = LK_CODCIA
 ser_llave!ser_fbg = Left(combotipo.text, 1)
 ser_llave!ser_numero = Val(tserie.text)
 ser_llave!ser_inicial = Val(tini.text)
 ser_llave!ser_final = Val(tfin.text)
 ser_llave!ser_tope = Val(tmax.text)
 ser_llave!ser_flag_ini = cheini.Value
 ser_llave!ser_activo = cheflag.Value
 ser_llave.Update
 LIMPIASERIES
End Sub

Private Sub numfac_b_f_GotFocus()
Azul numfac_b_f, numfac_b_f
End Sub

Private Sub numfac_b_f_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 serie_f.SetFocus
End If

End Sub

Private Sub numfac_b_GotFocus()
Azul numfac_b, numfac_b
End Sub

Private Sub numfac_b_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 numfac_b_f.SetFocus
End If

End Sub

Private Sub numfac_f_f_GotFocus()
Azul numfac_f_f, numfac_f_f
End Sub

Private Sub numfac_f_f_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii <> 13 Then
  Exit Sub
End If
If CmdModificar.Enabled Then
   CmdModificar.SetFocus
Else
   CmdAgregar.SetFocus
End If

End Sub

Private Sub numfac_f_GotFocus()
Azul numfac_f, numfac_f
End Sub

Private Sub numfac_f_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  numfac_f_f.SetFocus
End If
End Sub

Private Sub numfac_g_f_GotFocus()
Azul numfac_g_f, numfac_g_f
End Sub

Private Sub numfac_g_f_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 Serie_b.SetFocus
End If

End Sub

Private Sub numfac_g_GotFocus()
Azul numfac_g, numfac_g
End Sub

Private Sub numfac_g_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 numfac_g_f.SetFocus
End If

End Sub

Private Sub opdevice_Click(Index As Integer)
WF = opdevice(Index).Tag
'WF = Mid(par_llave!par_device_fbg, Index + 1, 1)
loc_key = 99
For fila = 0 To cmdimpresoras.ListCount - 1
  cmdimpresoras.ListIndex = fila
  If Trim(Right(cmdimpresoras.text, 4)) = WF Then
    Exit For
  End If
Next fila
loc_key = 0

End Sub

Private Sub par_moneda_con_KeyPress(KeyAscii As Integer)
WCAR = UCase(Chr(KeyAscii))
If WCAR = "S" Or WCAR = "D" Or KeyAscii = 8 Then
Else
  KeyAscii = 0
  Exit Sub
End If

End Sub

Private Sub Serie_b_GotFocus()
Azul Serie_b, Serie_b
End Sub

Private Sub Serie_b_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 numfac_b.SetFocus
End If

End Sub

Private Sub serie_f_GotFocus()
Azul serie_f, serie_f
End Sub

Private Sub serie_f_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 numfac_f.SetFocus
End If

End Sub

Private Sub serie_g_GotFocus()
Azul serie_g, serie_g
End Sub

Private Sub serie_g_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 numfac_g.SetFocus
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

Private Sub tfin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Azul tmax, tmax
End If

End Sub

Private Sub tini_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Azul tfin, tfin
End If
End Sub

Private Sub tmax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   grabaserie.SetFocus
End If


End Sub

Private Sub tserie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 PUB_FBG = Left(combotipo.text, 1)
 PUB_NUMSER = Val(tserie.text)
 LEER_SER_LLAVE
 If ser_llave.EOF Then
    pub_mensaje = "Serie No Existe. Desea Crear la nueva SERIE..?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
       Exit Sub
    End If
    Azul tini, tini
 Else
    For fila = 0 To combotipo.ListCount - 1
      If Trim(ser_llave!ser_fbg) = Left(combotipo.text, 1) Then Exit For
    Next fila
    tserie.text = ser_llave!ser_numero
    tini.text = ser_llave!ser_inicial
    tfin.text = ser_llave!ser_final
    tmax.text = ser_llave!ser_tope
    cheini.Value = Val(ser_llave!ser_flag_ini)
    cheflag.Value = Val(ser_llave!ser_activo)
    Azul tini, tini
 End If
End If

End Sub

Private Sub txt_key_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  If txtnombre.Enabled Then
    txtnombre.SetFocus
  End If
End If
End Sub

Private Sub txtcolor_Change()
On Error GoTo DE
lblcia.Caption = Trim(txtnombre.text)
lblcia.BackColor = QBColor(txtcolor.text)
lblcia.ForeColor = QBColor(txttexto.text)
DE:
End Sub

Private Sub txtcolor_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
End Sub

Private Sub txtmaxv_GotFocus()
Azul txtmaxv, txtmaxv
End Sub

Private Sub txtmaxv_KeyPress(KeyAscii As Integer)
If KeyAscii = 45 And Trim(txtmaxv.text) = "" Then
  Exit Sub
End If
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 serie_g.SetFocus
End If

End Sub

Private Sub txtmoneda_v_KeyPress(KeyAscii As Integer)
WCAR = UCase(Chr(KeyAscii))
If WCAR = "S" Or WCAR = "D" Or WCAR = "A" Or KeyAscii = 8 Then
Else
  KeyAscii = 0
  Exit Sub
End If
End Sub

Private Sub txtmorosos_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
End Sub

Private Sub txtnombre_GotFocus()
Azul txtnombre, txtnombre
End Sub

Private Sub txtnombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 txtnomcorto.SetFocus
End If
End Sub

Private Sub txtnomcorto_GotFocus()
Azul txtnomcorto, txtnomcorto
End Sub

Private Sub txtnomcorto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 txtmaxv.SetFocus
End If

End Sub

Private Sub txttexto_Change()
On Error GoTo DE
lblcia.Caption = Trim(txtnombre.text)
lblcia.BackColor = QBColor(txtcolor.text)
lblcia.ForeColor = QBColor(txttexto.text)
DE:

End Sub

Private Sub txttexto_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
End Sub

Private Sub txttipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Command1.SetFocus
Exit Sub
End If
SOLO_DECIMAL txttipo, KeyAscii

End Sub

Private Sub txttipo_LostFocus()
txttipo.text = Format(txttipo, "0.0000")
End Sub
Public Sub LLENADO_FAM()
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
        art_familia.AddItem tab_mayor!TAB_NOMLARGO & String(80, " ") & Trim(CStr(tab_mayor!TAB_NUMTAB))
        tab_mayor.MoveNext
    Loop
End Sub


Public Sub LIMPIASERIES()
    tserie.text = ""
    tini.text = ""
    tfin.text = ""
    tmax.text = ""
    cheini.Value = 0
    cheflag.Value = 0
End Sub
