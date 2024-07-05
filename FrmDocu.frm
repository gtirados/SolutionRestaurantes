VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmDocu 
   Caption         =   "Consulta de Operaciones"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12810
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   12810
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "Genera TXT SFS"
      Height          =   720
      Left            =   10560
      TabIndex        =   88
      Top             =   1440
      Width           =   990
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   8
      Top             =   -120
      Width           =   10380
      Begin VB.TextBox txtvend 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3045
         TabIndex        =   58
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdSiguiente 
         Height          =   435
         Left            =   9765
         Picture         =   "FrmDocu.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   200
         Width           =   435
      End
      Begin VB.CommandButton CmdAnterior 
         Height          =   435
         Left            =   9225
         Picture         =   "FrmDocu.frx":0942
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   200
         Width           =   450
      End
      Begin VB.ComboBox TIPMOV 
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
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtNumfac 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   6795
         MaxLength       =   9
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cmbFBG 
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
         ForeColor       =   &H00404040&
         Height          =   315
         ItemData        =   "FrmDocu.frx":1044
         Left            =   3720
         List            =   "FrmDocu.frx":1046
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtSerie 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   6270
         MaxLength       =   3
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblvend 
         Caption         =   "Vend."
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
         Left            =   3030
         TabIndex        =   59
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblNumfac 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Documento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   6240
         TabIndex        =   10
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6225
         TabIndex        =   53
         Top             =   330
         Width           =   1980
      End
      Begin VB.Label lbldocu 
         Alignment       =   2  'Center
         Caption         =   "Operación"
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
         Index           =   6
         Left            =   135
         TabIndex        =   42
         Top             =   135
         Width           =   2655
      End
      Begin VB.Label lbldocu 
         Alignment       =   2  'Center
         Caption         =   "Tipo"
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
         Index           =   0
         Left            =   3720
         TabIndex        =   9
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Frame FRADIRE 
      Caption         =   "Direccion de Entrega de Mercaderia :"
      Height          =   735
      Left            =   480
      TabIndex        =   69
      Top             =   3405
      Visible         =   0   'False
      Width           =   8055
      Begin VB.ComboBox TxtZonaTrabajo 
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
         Left            =   3960
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   360
         WhatsThisHelpID =   5
         Width           =   1935
      End
      Begin VB.ComboBox TxtSubZonaTrabajo 
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
         Left            =   6000
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   360
         WhatsThisHelpID =   6
         Width           =   1935
      End
      Begin VB.TextBox txtnum 
         Height          =   285
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   71
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtdire 
         Height          =   285
         Left            =   240
         MaxLength       =   30
         TabIndex        =   70
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.CheckBox imp 
      BackColor       =   &H008B4914&
      Caption         =   "&Directo a Impresora"
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
      Left            =   10560
      TabIndex        =   68
      Top             =   2640
      Width           =   1215
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
      Left            =   10560
      Picture         =   "FrmDocu.frx":1048
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   5085
      Width           =   1155
   End
   Begin VB.CommandButton cmdImp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Imprimir"
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
      Left            =   10560
      MaskColor       =   &H00FFFFFF&
      Picture         =   "FrmDocu.frx":18BE
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   3240
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Documento "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   6495
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   10395
      Begin VB.ComboBox cmbMotivo 
         Height          =   315
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   86
         Top             =   6120
         Width           =   2895
      End
      Begin VB.TextBox FECHA_PART 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9270
         TabIndex        =   77
         Top             =   4665
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox tguia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9270
         TabIndex        =   76
         Top             =   4095
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CheckBox cherela 
         BackColor       =   &H00808080&
         Caption         =   "Guia Rem."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   9240
         TabIndex        =   75
         Top             =   3795
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox NUMERO 
         Height          =   285
         Left            =   4560
         TabIndex        =   63
         Top             =   6075
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox sin_valor 
         Caption         =   "&Guia Sin Valor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8310
         TabIndex        =   61
         Top             =   1230
         Width           =   1275
      End
      Begin VB.ComboBox TRANS 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   6075
         Width           =   4335
      End
      Begin MSFlexGridLib.MSFlexGrid grid_fac2 
         Height          =   2775
         Left            =   75
         TabIndex        =   6
         Top             =   2160
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   4895
         _Version        =   393216
         ForeColor       =   4210752
         BackColorFixed  =   12632256
         ForeColorFixed  =   4210752
         ForeColorSel    =   16511715
         BackColorBkg    =   16777215
         GridColorFixed  =   8421504
         Enabled         =   -1  'True
         HighLight       =   2
         GridLinesFixed  =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.Frame Frame3 
         Height          =   810
         Left            =   60
         TabIndex        =   14
         Tag             =   "119"
         Top             =   4980
         Width           =   8835
         Begin MSComctlLib.StatusBar stbEtiqueta 
            Height          =   330
            Left            =   30
            TabIndex        =   74
            Top             =   120
            Width           =   8745
            _ExtentX        =   15425
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
               NumPanels       =   6
               BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  Bevel           =   2
                  Text            =   "SubTotal"
                  TextSave        =   "SubTotal"
               EndProperty
               BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  Bevel           =   2
                  Text            =   "Dsto.Gral."
                  TextSave        =   "Dsto.Gral."
               EndProperty
               BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  Bevel           =   2
                  Text            =   "Dsctos."
                  TextSave        =   "Dsctos."
               EndProperty
               BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  Bevel           =   2
                  Text            =   "Impto."
                  TextSave        =   "Impto."
               EndProperty
               BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  Bevel           =   2
               EndProperty
               BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  Bevel           =   2
                  Text            =   "Total"
                  TextSave        =   "Total"
               EndProperty
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label d_flete 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5880
            TabIndex        =   31
            ToolTipText     =   "Doble Click para modificar..."
            Top             =   465
            Width           =   1440
         End
         Begin VB.Label d_neto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7335
            TabIndex        =   30
            Top             =   465
            Width           =   1440
         End
         Begin VB.Label d_descto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2955
            TabIndex        =   29
            Top             =   465
            Width           =   1440
         End
         Begin VB.Label d_impto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4410
            TabIndex        =   28
            Top             =   465
            Width           =   1440
         End
         Begin VB.Label d_gastos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1500
            TabIndex        =   27
            Top             =   465
            Width           =   1440
         End
         Begin VB.Label d_subtotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   30
            TabIndex        =   26
            Top             =   465
            Width           =   1440
         End
         Begin VB.Label lblflete 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   4920
            TabIndex        =   20
            Tag             =   "9999"
            Top             =   120
            Width           =   1005
         End
         Begin VB.Label LCODART 
            Caption         =   "TOTAL"
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
            Index           =   2
            Left            =   7560
            TabIndex        =   19
            Tag             =   "9999"
            Top             =   120
            Width           =   765
         End
         Begin VB.Label LCODART 
            Caption         =   "Impto."
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
            Index           =   4
            Left            =   3960
            TabIndex        =   18
            Tag             =   "9999"
            Top             =   120
            Width           =   525
         End
         Begin VB.Label LCODART 
            Caption         =   "Descto."
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
            Index           =   5
            Left            =   2760
            TabIndex        =   17
            Tag             =   "9999"
            Top             =   120
            Width           =   765
         End
         Begin VB.Label LCODART 
            Caption         =   "Gastos"
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
            Index           =   6
            Left            =   1560
            TabIndex        =   16
            Tag             =   "9999"
            Top             =   120
            Width           =   765
         End
         Begin VB.Label LCODART 
            Caption         =   "Subtotal:"
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
            Index           =   7
            Left            =   240
            TabIndex        =   15
            Tag             =   "9999"
            Top             =   120
            Width           =   765
         End
      End
      Begin VB.CheckBox chetrans 
         Caption         =   "Imprimir Datos del Transportista en Guia"
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
         TabIndex        =   55
         Top             =   5835
         Visible         =   0   'False
         Width           =   4215
      End
      Begin ComctlLib.ProgressBar PB 
         Height          =   165
         Left            =   3480
         TabIndex        =   32
         Top             =   360
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   291
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Motivo de Traslado"
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
         Left            =   7320
         TabIndex        =   87
         Top             =   5880
         Width           =   1365
      End
      Begin VB.Label lbldocu 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Part."
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
         Index           =   2
         Left            =   9240
         TabIndex        =   84
         Top             =   4425
         Width           =   975
      End
      Begin VB.Label d_dias 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9885
         TabIndex        =   83
         Top             =   2265
         Width           =   375
      End
      Begin VB.Label d_newvcto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9270
         TabIndex        =   82
         Top             =   3435
         Width           =   885
      End
      Begin VB.Label lbldocu 
         BackStyle       =   0  'Transparent
         Caption         =   "Dias:"
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
         Index           =   3
         Left            =   9360
         TabIndex        =   81
         Top             =   2265
         Width           =   495
      End
      Begin VB.Label lbldocu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N.Vcto."
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
         Index           =   1
         Left            =   9480
         TabIndex        =   80
         Top             =   3210
         Width           =   570
      End
      Begin VB.Label d_fechaV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9240
         TabIndex        =   79
         Top             =   2865
         Width           =   1005
      End
      Begin VB.Label lbldocu 
         BackStyle       =   0  'Transparent
         Caption         =   " Vcto."
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
         Index           =   5
         Left            =   9360
         TabIndex        =   78
         Top             =   2595
         Width           =   735
      End
      Begin VB.Label FCONT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fec. Contable :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   120
         TabIndex        =   64
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label d_nomven 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7455
         TabIndex        =   62
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label d_codven 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6960
         TabIndex        =   22
         Top             =   840
         Width           =   405
      End
      Begin VB.Label d_fecha_compra 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1620
         TabIndex        =   56
         Top             =   525
         Width           =   1155
      End
      Begin VB.Label d_fecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/01/2001"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1620
         TabIndex        =   24
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label l_fecha_compra 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Emisión :"
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
         Height          =   195
         Left            =   120
         TabIndex        =   57
         Top             =   525
         Width           =   1365
      End
      Begin VB.Label d_moneda 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5160
         TabIndex        =   43
         Top             =   6075
         Width           =   375
      End
      Begin VB.Label txtdocu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   8295
         TabIndex        =   44
         Top             =   1845
         Width           =   1935
      End
      Begin VB.Label d_dire 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1830
         TabIndex        =   36
         Top             =   1860
         Width           =   6345
      End
      Begin VB.Label d_usuario 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   52
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lbldireccion 
         Alignment       =   1  'Right Justify
         Caption         =   "Dir. Entrega :"
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
         Left            =   105
         TabIndex        =   38
         Top             =   1875
         Width           =   1665
      End
      Begin VB.Label lblfac 
         Caption         =   "Documento"
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
         Height          =   225
         Left            =   8295
         TabIndex        =   45
         Top             =   1575
         Width           =   1095
      End
      Begin VB.Label lbldomicilio 
         Alignment       =   1  'Right Justify
         Caption         =   "Domicilio :"
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
         Left            =   30
         TabIndex        =   51
         Top             =   1545
         Width           =   975
      End
      Begin VB.Label d_domicilio 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   1020
         TabIndex        =   50
         Top             =   1545
         Width           =   7155
      End
      Begin VB.Label d_efectivo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   210
         TabIndex        =   49
         Tag             =   "9999"
         Top             =   6075
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblEfectivo 
         Caption         =   "Total Efectivo ."
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
         Left            =   240
         TabIndex        =   48
         Tag             =   "9999"
         Top             =   5835
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblsaldo 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Actual"
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
         Left            =   5760
         TabIndex        =   34
         Tag             =   "9999"
         Top             =   5835
         Width           =   1275
      End
      Begin VB.Label lblcheque 
         Caption         =   "Total Cheque ."
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
         Left            =   2280
         TabIndex        =   47
         Tag             =   "9999"
         Top             =   5835
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label d_cheque 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2250
         TabIndex        =   46
         Tag             =   "9999"
         Top             =   6075
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label d_mensaje 
         Alignment       =   2  'Center
         Caption         =   "DOCUMENTO EN BLANCO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   41
         Top             =   600
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label d_ruc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   4605
         TabIndex        =   40
         Top             =   870
         Width           =   1215
         WordWrap        =   -1  'True
      End
      Begin VB.Label LBLRUC 
         Caption         =   "R.U.C."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5040
         TabIndex        =   39
         Top             =   600
         Width           =   495
      End
      Begin VB.Label d_saldo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   5640
         TabIndex        =   35
         Tag             =   "9999"
         Top             =   6075
         Width           =   1485
      End
      Begin VB.Label LBLEXTORNO 
         Caption         =   "DOCUMENTO EXTORNADO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3240
         TabIndex        =   33
         Top             =   360
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label lblcondicion 
         Caption         =   "Condición"
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
         Left            =   6585
         TabIndex        =   23
         Top             =   150
         Width           =   960
      End
      Begin VB.Label d_condicion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7560
         TabIndex        =   13
         Top             =   135
         Width           =   2565
      End
      Begin VB.Label d_nomclie 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   1830
         TabIndex        =   12
         Top             =   1230
         Width           =   6345
      End
      Begin VB.Label d_Codclie 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   720
         TabIndex        =   11
         Top             =   1230
         Width           =   1005
      End
      Begin VB.Label lblpersona 
         Caption         =   "Cliente"
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
         Left            =   105
         TabIndex        =   21
         Top             =   1245
         Width           =   600
      End
      Begin VB.Label lbldocu 
         Caption         =   "Fecha Proceso :"
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
         Index           =   7
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Height          =   2835
         Left            =   8895
         TabIndex        =   85
         Top             =   2175
         Width           =   1410
      End
   End
   Begin VB.TextBox tmoneda 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
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
      Left            =   6210
      TabIndex        =   54
      Text            =   "S/."
      Top             =   4920
      Width           =   495
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   240
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4914&
      BorderStyle     =   1  'Fixed Single
      Height          =   7095
      Index           =   5
      Left            =   10440
      TabIndex        =   65
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmDocu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DOC102 As String
Dim loc_flag_espera As String * 1
Dim pstransporte As rdoQuery
Dim TRANSPORTE As rdoResultset
Dim PSFAR As rdoQuery
Dim far_r As rdoResultset
Dim PSFAR_CONSUL As rdoQuery
Dim far_consul As rdoResultset
Dim PSCAR_CONSUL As rdoQuery
Dim car_consul As rdoResultset
Dim wflag_docu As String * 1
Dim temporal
Dim tempo_serie
Dim LOC_TIPMOV As Integer
Dim LOC_NUMFAC_FIN As Currency
Dim WGUIA_RELA As String
Dim SERIE_GUIA As String * 3
Dim PS_VE2 As rdoQuery
Dim VE2_LLAVE As rdoResultset
Dim SIN_CODART As Integer
Dim PS_TRA As rdoQuery
Dim llave_trans As rdoResultset
Dim LOC_ARROZ As String * 1
'variables agregadas
Dim rs As rdoResultset
Dim PS As rdoQuery





Private Sub cherela_Click()
 If tguia.Visible Then tguia.SetFocus
 If cherela.Value = 1 Then
   If val(WGUIA_RELA) <> 0 Then
      tguia.Text = "G/." + Trim(SERIE_GUIA) + " - " + WGUIA_RELA
   End If
 End If
' If LK_EMP = "HER" Then
 '   If cherela.Value = 1 Then
 '      tguia.Text = "G/." + Trim(usu_llave!USU_SERIE_G) + " - "
 '      tguia.SetFocus
 '      tguia.SelStart = Len(tguia.Text)
 '   Else
 '     tguia.Text = ""
 '   End If
 'End If
 NUMERO.Visible = True
End Sub

Private Sub chetrans_Click()
If chetrans.Value = 0 Then
  TRANS.ListIndex = -1
  TRANS.Enabled = False
Else
  TRANS.Enabled = True
End If
If TRANS.Visible Then
 If TRANS.ListCount <> 0 Then
   TRANS.ListIndex = TRANS.ListIndex
 End If
  If TRANS.Enabled Then TRANS.SetFocus
End If
End Sub

Private Sub cmbFBG_Click()
If temporal = "X" Then
 Exit Sub
End If
If LK_FLAG_FACTURACION = "V" Then
 If val(txtvend.Text) <> 0 Then
 If ven_llave.EOF Then
   MsgBox "Digite un Vendedor", 48, Pub_Titulo
   Exit Sub
 End If
 Else
   Exit Sub
 End If
End If

If LOC_TIPMOV = 97 And (Right(cmbFBG.Text, 1) = "P" Or Right(cmbFBG.Text, 1) = "C") Then
  pub_cadena = "SELECT * FROM facart WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_NUMSER = ? AND FAR_FBG=? AND FAR_NUMFAC = ? AND FAR_CP = ?   ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_NUMSER, FAR_FBG, FAR_NUMFAC, FAR_NUMSEC"
ElseIf LOC_TIPMOV = 99 Then
  pub_cadena = "SELECT * FROM facart WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_NUMSER = ? AND FAR_FBG=? AND FAR_NUMFAC = ?   ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_NUMSER, FAR_NUMFAC, FAR_NUMSEC"
Else
  pub_cadena = "SELECT * FROM facart WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_NUMSER = ? AND FAR_FBG=? AND FAR_NUMFAC = ? AND FAR_VISIBLE = 1   ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_NUMSER, FAR_FBG, FAR_NUMFAC, FAR_NUMSEC"
End If
Set PSFAR = CN.CreateQuery("", pub_cadena)
 PSFAR.rdoParameters(0) = 0
 PSFAR.rdoParameters(1) = " "
 PSFAR.rdoParameters(2) = 0
 PSFAR.rdoParameters(3) = " "
 PSFAR.rdoParameters(4) = 0
 If LOC_TIPMOV = 97 And (Right(cmbFBG.Text, 1) = "P" Or Right(cmbFBG.Text, 1) = "C") Then
  PSFAR.rdoParameters(5) = " "
 End If
Set far_r = PSFAR.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

LIMPIA_DOCU
grid_fac2.rows = 2
txtNumfac.Text = ""
txtSerie.Text = ""
'SQ_OPER = 1
'PUB_CODCIA = LK_CODCIA
'LEER_PAR_LLAVE
'If par_llave.EOF Then
'  Exit Sub
'End If
txtSerie.Locked = False
lblsaldo.Caption = "Saldo Actual "
lbldomicilio.Caption = "Domicilio :"
lbldocu(1).Visible = True
lbldocu(3).Visible = True
lbldocu(5).Visible = True
d_dias.Visible = True
d_fechaV.Visible = True
d_newvcto.Visible = True

'If Left(cmbFBG.Text, 1) = "P" Then
' lblcheque.Visible = True
' d_cheque.Visible = True
' d_efectivo.Visible = True
' lblEfectivo.Visible = True
' lblSaldo.Caption = "Total Planilla="
 'txtSerie.Locked = True
' lbldomicilio.Caption = "Vendedor :"
' txtSerie.Text = 0
' lbldocu(1).Visible = False
' lbldocu(3).Visible = False
' lbldocu(5).Visible = False
' d_dias.Visible = False
' d_fechaV.Visible = False
' d_newvcto.Visible = False
' Exit Sub
'End If
If LK_FLAG_FACTURACION = "V" Then
   Select Case Left(cmbFBG.Text, 1)
   Case "G"
       txtSerie.Text = ven_llave!VEM_SERIE_G
   Case "B"
       txtSerie.Text = ven_llave!VEM_SERIE_B
   Case "F"
       txtSerie.Text = ven_llave!VEM_SERIE_F
   Case "P"
       txtSerie.Text = ven_llave!VEM_SERIE_P
   End Select
ElseIf LK_FLAG_FACTURACION = "A" Then
   If Left(cmbFBG.Text, 1) = "F" Then
     txtSerie.Text = par_llave!PAR_F_SERIE
   ElseIf Left(cmbFBG.Text, 1) = "B" Then
     txtSerie.Text = par_llave!PAR_B_SERIE
   ElseIf Left(cmbFBG.Text, 1) = "G" Then
     txtSerie.Text = par_llave!PAR_G_SERIE
   ElseIf Left(cmbFBG.Text, 1) = "T" Then
     txtSerie.Text = 1
   ElseIf Left(cmbFBG.Text, 1) = "P" Then
     txtSerie.Text = 1
   End If
ElseIf LK_FLAG_FACTURACION = "U" Then
   If Left(cmbFBG.Text, 1) = "F" Then
     txtSerie.Text = usu_llave!USU_SERIE_F
   ElseIf Left(cmbFBG.Text, 1) = "B" Then
     txtSerie.Text = usu_llave!USU_SERIE_B
   End If
End If
If Left(cmbFBG.Text, 1) = "N" Or Left(cmbFBG.Text, 1) = "D" Then
     If Left(cmbFBG.Text, 1) = "D" Then
        txtSerie.Text = par_llave!PAR_SERIE_NDEB
     Else
        txtSerie.Text = par_llave!PAR_SERIE_NCRE
     End If
    If LK_FLAG_FACTURACION = "U" Then
      If Left(cmbFBG.Text, 1) = "D" Then
        txtSerie.Text = usu_llave!USU_SERIE_ND
      Else
        txtSerie.Text = usu_llave!USU_SERIE_NC
      End If
    End If
    lbldocu(1).Visible = False
    lbldocu(3).Visible = False
    lbldocu(5).Visible = False
    d_dias.Visible = False
    d_fechaV.Visible = False
    d_newvcto.Visible = False
End If
If LOC_TIPMOV = 10 Then
 PU_TIPMOV = 10
 PU_NUMSER = val(txtSerie.Text)
 PU_FBG = Left(cmbFBG.Text, 1)
ElseIf LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
 PU_TIPMOV = LOC_TIPMOV
 PU_NUMSER = val(txtSerie.Text)
 PU_FBG = Left(cmbFBG.Text, 1)
ElseIf LOC_TIPMOV = 20 Or LOC_TIPMOV = 99 Then
 PU_TIPMOV = LOC_TIPMOV
 If LOC_TIPMOV = 20 Then
  txtSerie.Text = val(par_llave!PAR_SER_KARDEX)
 End If
 PU_NUMSER = val(txtSerie.Text)
 If LOC_TIPMOV = 99 Then
   PU_FBG = "K"
 Else
   PU_FBG = " "
 End If

Else
If LOC_TIPMOV = 5 Then
   PSCNT_LLAVE.rdoParameters(2) = 1
ElseIf LOC_TIPMOV = 6 Then
   PSCNT_LLAVE.rdoParameters(2) = 0
End If

If LOC_TIPMOV = 5 Or LOC_TIPMOV = 6 Then
   PSCNT_LLAVE.rdoParameters(0) = LK_CODCIA
   PSCNT_LLAVE.rdoParameters(1) = 2403
   cnt_llave.Requery
   If Not cnt_llave.EOF Then txtSerie.Text = Nulo_Valor0(cnt_llave!cnt_serie)
Else
 txtSerie.Text = "0"
End If
If LOC_TIPMOV = 3 Then pu_cp = "P"
If LOC_TIPMOV = 102 Then pu_cp = "C"
If LOC_TIPMOV = 100 Or LOC_TIPMOV = 101 Or LOC_TIPMOV = 93 Then
If LOC_TIPMOV = 100 Then pu_cp = "C"
 txtSerie.Text = "1"
End If

 PU_TIPMOV = LOC_TIPMOV
 PU_NUMSER = val(txtSerie.Text)
 PU_FBG = " "

End If
pu_codcia = LK_CODCIA
LEER_FAR_CONSUL
If Not far_consul.EOF Then
 txtNumfac.Text = far_consul!far_numfac
Else
 txtNumfac.Text = "0"
End If
Azul txtNumfac, txtNumfac

If Trim(d_fecha.Caption) = "" Then txtnumfac_KeyPress 13

Exit Sub


End Sub

Private Sub cmbFBG_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
 Exit Sub
End If
If temporal = "X" Then
 Exit Sub
End If
'SQ_OPER = 1
'PUB_CODCIA = LK_CODCIA
'LEER_PAR_LLAVE
'If par_llave.EOF Then
'  Exit Sub
'End If
If LOC_TIPMOV <> 3 Then pu_cp = " "
If Left(cmbFBG.Text, 1) = "P" And PUB_TIPMOV <> 10 Then
 txtNumfac.Text = Nulo_Valor0(par_llave!par_planilla)
 txtNumfac.SetFocus
 txtnumfac_KeyPress 13
 Exit Sub
End If
If LOC_TIPMOV = 30 Then
 Dim PSTEMP_MAYOR As rdoQuery
 Dim temp_mayor  As rdoResultset
 Dim wser
 Dim wnumfac As Currency
 pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_CODCIA = ?  and ped_tipmov=20  ORDER BY  PED_NUMFAC DESC "
 Set PSTEMP_MAYOR = CN.CreateQuery("", pub_cadena)
 PSTEMP_MAYOR.rdoParameters(0) = " "
 PSTEMP_MAYOR.MaxRows = 1
 Set temp_mayor = PSTEMP_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
 PSTEMP_MAYOR(0) = LK_CODCIA
 temp_mayor.Requery
 If temp_mayor.EOF Then
    wser = 0
    wnumfac = 0
 Else
    wser = Nulo_Valors(temp_mayor!PED_numser)
    wnumfac = Nulo_Valor0(temp_mayor!PED_NUMFAC)
 End If
 pu_cp = "P"
 txtSerie.Text = wser
 txtNumfac.Text = wnumfac
 txtNumfac.SetFocus
 txtnumfac_KeyPress 13
 Exit Sub
End If

If LK_FLAG_FACTURACION = "V" Then
   If ven_llave.EOF Then GoTo dale
   Select Case Left(cmbFBG.Text, 1)
   Case "G"
       txtSerie.Text = ven_llave!VEM_SERIE_G
   Case "B"
       txtSerie.Text = ven_llave!VEM_SERIE_B
   Case "F"
       txtSerie.Text = ven_llave!VEM_SERIE_F
   Case "P"
       txtSerie.Text = ven_llave!VEM_SERIE_P
   End Select
dale:
ElseIf LK_FLAG_FACTURACION = "A" And LOC_TIPMOV <> 93 Then
 If Left(cmbFBG.Text, 1) = "F" Then
  txtSerie.Text = par_llave!PAR_F_SERIE
 ElseIf Left(cmbFBG.Text, 1) = "B" Then
  txtSerie.Text = par_llave!PAR_B_SERIE
 ElseIf Left(cmbFBG.Text, 1) = "G" Then
  txtSerie.Text = par_llave!PAR_G_SERIE
 ElseIf Left(cmbFBG.Text, 1) = "N" Then
  txtSerie.Text = par_llave!PAR_SERIE_NCRE
 ElseIf Left(cmbFBG.Text, 1) = "D" Then
  txtSerie.Text = par_llave!PAR_SERIE_NDEB
 End If
ElseIf LK_FLAG_FACTURACION = "U" Then
   If Left(cmbFBG.Text, 1) = "F" Then
     txtSerie.Text = usu_llave!USU_SERIE_F
   ElseIf Left(cmbFBG.Text, 1) = "B" Then
     txtSerie.Text = usu_llave!USU_SERIE_B
   End If
End If
If LOC_TIPMOV = 10 Then
 pu_cp = "C"
 PU_TIPMOV = 10
 PU_NUMSER = val(txtSerie.Text)
 PU_FBG = Left(cmbFBG.Text, 1)
ElseIf LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
 PU_TIPMOV = LOC_TIPMOV
 PU_NUMSER = val(txtSerie.Text)
 PU_FBG = Left(cmbFBG.Text, 1)
 pu_cp = "C"
 If Right(Trim(cmbFBG.Text), 1) = "P" Then
   txtSerie.Text = "0"
   PU_NUMSER = 0
   PU_FBG = "C"
   pu_cp = "P"
 End If
 
ElseIf LOC_TIPMOV = 20 Or LOC_TIPMOV = 99 Then
 pu_cp = "P"
 PU_TIPMOV = LOC_TIPMOV
 PU_NUMSER = 0
 PU_FBG = " "
 If LOC_TIPMOV = 20 And Left(cmbFBG.Text, 1) <> "K" Then
    If Left(cmbFBG.Text, 1) = "F" Then
      pub_cadena = "SELECT * FROM facart WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_NUMSER_C = ? AND FAR_NUMFAC_C = ? ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_NUMSER, FAR_FBG, FAR_NUMFAC, FAR_NUMSEC"
    Else
      pub_cadena = "SELECT * FROM facart WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_NUMGUIA = ?   ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_NUMSER, FAR_FBG, FAR_NUMFAC, FAR_NUMSEC"
    End If
    Set PSFAR = CN.CreateQuery("", pub_cadena)
    If Left(cmbFBG.Text, 1) = "F" Then
        PSFAR.rdoParameters(0) = 0
        PSFAR.rdoParameters(1) = 0
        PSFAR.rdoParameters(2) = 0
        PSFAR.rdoParameters(3) = 0
    Else
        PSFAR.rdoParameters(0) = 0
        PSFAR.rdoParameters(1) = 0
        PSFAR.rdoParameters(2) = 0
    End If
    Set far_r = PSFAR.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
 End If
 If LOC_TIPMOV = 99 Then
  PU_FBG = Left(cmbFBG.Text, 1)
 End If
Else
'' txtserie.Text = "0"
 PU_TIPMOV = LOC_TIPMOV
 PU_NUMSER = 0
 PU_FBG = " "
End If
PU_NUMSER = val(txtSerie.Text)
pu_codcia = LK_CODCIA
LEER_FAR_CONSUL
If Not far_consul.EOF Then
 txtNumfac.Text = far_consul!far_numfac
Else
 txtNumfac.Text = "0"
End If
txtNumfac.SetFocus
txtnumfac_KeyPress 13
If LK_EMP = "HER" And val(txtSerie.Text) = 0 And LOC_TIPMOV = 10 Then
   txtSerie.Locked = False
   txtSerie.Text = ""
   txtSerie.SetFocus
End If
End Sub

Private Sub cmdAnterior_Click()
Dim tempo
If LOC_TIPMOV = 0 Then Exit Sub
tempo = val(txtNumfac.Text)
If LOC_TIPMOV = 10 Then
 If Trim(txtSerie.Text) = "" Then
  Exit Sub
 End If
End If
If val(txtNumfac.Text) <= 0 Then
 LIMPIA_DOCU
 grid_fac2.Clear
 Exit Sub
End If
txtNumfac.Text = val(txtNumfac.Text) - 1
If LOC_TIPMOV = 96 Or LOC_TIPMOV = 30 Then ' PLANILLA
 txtnumfac_KeyPress 13
 Exit Sub
End If
wflag_docu = ""
loc_flag_espera = "A"
LLENA_CONSULTA
loc_flag_espera = ""
If wflag_docu = "A" Then
  If Trim(d_fecha.Caption) <> "" Then LIMPIA_DOCU
  d_mensaje.Visible = True
  'CmdAnterior.Enabled = False
  'Beep
  'Beep
Else
  d_mensaje.Visible = False
  CmdAnterior.Enabled = True
End If
Azul txtNumfac, txtNumfac
LOC_NUMFAC_FIN = val(txtNumfac.Text)

End Sub

Private Sub cmdCerrar_Click()
Unload frmdocu
End Sub

Private Sub cmdCommand1_Click()
'       If Me.DatTiposDoctos.BoundText = "01" Or Me.DatTiposDoctos.BoundText = "03" Then
'           CrearArchivoPlano Left(Me.DatTiposDoctos.Text, 1), Me.lblSerie.Caption, Me.txtNro.Text
'       End If
       On Error GoTo Exporta
        If LOC_TIPMOV = 10 Or LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
            CrearArchivoPlano2 Left(cmbFBG.Text, 1), Me.txtSerie.Text, Me.txtNumfac.Text
        End If
        MsgBox "Archivos generados correctamente.", vbInformation, Pub_Titulo
       Exit Sub
Exporta:
       MsgBox "Error al exportar el archivo Plano", vbCritical, Pub_Titulo
       
       
End Sub

Private Sub cmdimp_Click()
If loc_flag_espera = "A" Then
 MsgBox "Espere ....!!!", 48, Pub_Titulo
 Exit Sub
End If
If LOC_TIPMOV = 0 Or Trim(d_fecha.Caption) = "" Then
Exit Sub
End If
If frmdocu.LBLEXTORNO.Visible Then
  MsgBox "Impresión No Procede...", 48, Pub_Titulo
  Exit Sub
End If
If LOC_TIPMOV = 30 Then
  GoSub ORDEN
  Exit Sub
End If
If chetrans.Value = 1 Then
 If TRANS.Text = "" Then
    MsgBox "Datos del Transportista: Seleccione uno de la lista", 48, Pub_Titulo
    Exit Sub
 End If
End If
SIN_CODART = 0
If LOC_TIPMOV = 75 Or LOC_TIPMOV = 99 Or LOC_TIPMOV = 102 Or LOC_TIPMOV = 3 Or LOC_TIPMOV = 100 Or LOC_TIPMOV = 101 Or LOC_TIPMOV = 93 Or LOC_TIPMOV = 20 Or LOC_TIPMOV = 5 Or LOC_TIPMOV = 6 Or LOC_TIPMOV = 10 Or LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
  If grid_fac2.TextMatrix(1, 1) = "" Then SIN_CODART = 1
  fila = REP_CONSUL
 Exit Sub
End If

If LOC_TIPMOV <> 96 Then
 Exit Sub
End If
If LOC_TIPMOV = 10 Then
 If chetrans.Value = 1 And val(Right(TRANS.Text, 3)) = 0 Then
      MsgBox "Está Activada a opción de Transportista. Seleccione un Transportista ?", 48, Pub_Titulo
     Exit Sub
 End If
End If
Dim i, j
Dim wranF
Dim LETRAS(24) As String * 1


Dim xl As Object
On Error GoTo FINTODO
Screen.MousePointer = 11
PB.Visible = True
PB.max = 6
PB.Min = 0
PB.Value = 0
PB.Value = PB.Value + 1
GoSub WEXCEL
PB.Value = PB.Value + 1
pub_cadena = ""
xl.Cells(4, 1) = "PLANILLA : " & Trim(txtSerie.Text) & " - " & Trim(txtNumfac.Text)
xl.Cells(3, 1) = "'" & Format(LK_FECHA_DIA, "dd/mm/yyyy")
xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
xl.Cells(2, 1) = "PLANILLA DE COBRANZA"
PB.Value = PB.Value + 1
For i = 0 To grid_fac2.rows - 1
  For j = 0 To 14
     If grid_fac2.TextMatrix(i, j) = "" Then
       xl.Cells(i + 7, j + 1) = " "
     Else
       xl.Cells(i + 7, j + 1) = grid_fac2.TextMatrix(i, j)
     End If
  Next j
Next i
PB.Value = PB.Value + 1
GoSub LETRAS
PB.Value = PB.Value + 1
wranF = "A" & i + 8 & ":D" & i + 8
xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 3
xl.Cells(i + 1 + 7, 1) = "Total Cheque.  ="
xl.Cells(i + 1 + 7, 2) = "'" & d_cheque.Caption
xl.Cells(i + 2 + 7, 1) = "Total Efectivo.="
xl.Cells(i + 2 + 7, 2) = "'" & d_efectivo.Caption
xl.Cells(i + 3 + 7, 1) = "Total Planilla.="
xl.Cells(i + 3 + 7, 2) = "'" & d_saldo.Caption

wranF = "A8:" & "O8"
xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 3
xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
xl.Cells(2, 1) = "PLANILA DE COBRANZA"
xl.Cells(3, 1) = "'" & Format(LK_FECHA_DIA, "dd/mm/yyyy")
PB.Value = PB.Value + 1
xl.DisplayAlerts = False
xl.Worksheets(1).Protect PUB_CLAVE
xl.APPLICATION.Visible = True
Set xl = Nothing
PB.Visible = False
Screen.MousePointer = 0
Exit Sub

WEXCEL:
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\CONSPLA.xls", 0, True, 4, PUB_CLAVE, PUB_CLAVE
Return

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

ORDEN:
Dim wser As String * 3
Dim WSRUTA As String
Dim wRuta As String
Dim rmoneda As String * 1
wRuta = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\"
frmdocu.Reportes.Connect = PUB_ODBC
frmdocu.Reportes.Destination = crptToWindow  '= crptToPrinter
frmdocu.Reportes.WindowLeft = 2
frmdocu.Reportes.WindowTop = 70
frmdocu.Reportes.WindowWidth = 635
frmdocu.Reportes.WindowHeight = 390
frmdocu.Reportes.Formulas(1) = ""
PUB_NETO = val(frmdocu.d_neto.Caption)
PUB_FECHA = frmdocu.d_fecha.Caption
PU_NUMSER = val((frmdocu.txtSerie.Text))

If Left(d_moneda.Caption, 3) = "US$" Then
   rmoneda = "D"
Else
   rmoneda = "S"
End If

PU_NUMFAC = val((frmdocu.txtNumfac.Text))
frmdocu.Reportes.Formulas(1) = "SON_EFECTIVO=  'SON: " & CONVER_LETRAS(PUB_NETO, rmoneda) & "'"
frmdocu.Reportes.WindowTitle = "ORDEN DE COMPRA  :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
frmdocu.Reportes.ReportFileName = wRuta + "ORDEN.RPT"
wser = PU_NUMSER
pub_cadena = "{PEDIDOS.PED_ESTADO} = 'N' AND {PEDIDOS.PED_CODCIA} = '" & LK_CODCIA & "' AND {PEDIDOS.PED_NUMSER}= '" & wser & "' AND {PEDIDOS.PED_NUMFAC} = " & PU_NUMFAC
frmdocu.Reportes.SelectionFormula = pub_cadena
frmdocu.Reportes.WindowTitle = frmdocu.Reportes.WindowTitle & " Archivo: " & Trim(frmdocu.Reportes.ReportFileName)
On Error GoTo accion
frmdocu.Reportes.Action = 1
On Error GoTo 0
Return
FINTODO:
accion:
 MsgBox Err.Description, 48, Pub_Titulo
 MsgBox "Reintente Nuevamente ..", 48, Pub_Titulo
End Sub

Private Sub cmdserie_Click()
Dim valor
valor = InputBox("Ingrese hasta que numero de " & Trim(cmbFBG.Text) & " Desea Mostrar para la Impresión. Segun serie : " & txtSerie.Text & " - ", "Inpresión en Serie . . . ", Trim(txtNumfac.Text))
If valor = "" Then Exit Sub
If val(valor) < val(txtNumfac.Text) Then
  MsgBox "No Procede... No puede ser menor que el Nº inicial ", 48, Pub_Titulo
  Exit Sub
End If
LOC_NUMFAC_FIN = valor
REP_CONSUL
LOC_NUMFAC_FIN = 0

End Sub

Private Sub cmdSiguiente_Click()
Dim tempo
If LOC_TIPMOV = 0 Then Exit Sub
tempo = val(txtNumfac.Text)
If LOC_TIPMOV = 10 Then
 If Trim(txtSerie.Text) = "" Then
  Exit Sub
 End If
End If
If val(txtNumfac.Text) < 0 Then
  Exit Sub
End If
txtNumfac.Text = val(txtNumfac.Text) + 1
If LOC_TIPMOV = 96 Or LOC_TIPMOV = 30 Then ' PLANILLA
 txtnumfac_KeyPress 13
 Exit Sub
End If
wflag_docu = ""
loc_flag_espera = "A"
LLENA_CONSULTA
loc_flag_espera = ""
If wflag_docu = "A" Then
  If Trim(d_fecha.Caption) <> "" Then LIMPIA_DOCU
  d_mensaje.Visible = True
Else
  d_mensaje.Visible = False
  cmdSiguiente.Enabled = True
End If
Azul txtNumfac, txtNumfac
LOC_NUMFAC_FIN = val(txtNumfac.Text)
End Sub

Private Sub d_descto_DblClick()
Dim cap_valor
Dim wcanti  As Currency
Dim wpeso  As Currency
If LK_EMP <> "3AA" Then Exit Sub
If LOC_TIPMOV <> 20 Then Exit Sub

 cap_valor = InputBox("Modificación de Descto de Mercaderia  en valor Porcentual(%).= " & Chr(13) & "el valor de descto. afecta a costo promedio mas no al documento.", " Descto(%)", d_descto.Caption)
 If cap_valor = "" Then Exit Sub
 If val(cap_valor) = 0 Then
  pub_mensaje = "Valor 0.00(%) para el Descto... desea continuar... "
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta <> vbYes Then
   Exit Sub
  End If
 End If
 
 wcanti = 0
 wpeso = 0
 fila = 1
 Do Until fila = grid_fac2.rows
     If Trim(grid_fac2.TextMatrix(fila, 9)) <> "NOT" Then
       grid_fac2.TextMatrix(fila, 6) = Format(val(grid_fac2.TextMatrix(fila, 5)) * (val(cap_valor) / 100), "0.00")
     Else
       grid_fac2.TextMatrix(fila, 6) = "0"
     End If
     ww_desc = ww_desc + val(grid_fac2.TextMatrix(fila, 6))
     fila = fila + 1
 Loop
pub_mensaje = "Chequear los datos del calculo. Total de Descto en " & d_moneda.Caption & " = " & Format(ww_desc, "0.00") & " - Confirmar la modificación ? "
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   txtnumfac_KeyPress 13
   Exit Sub
End If
   


pub_cadena = "SELECT FAR_PORDESCTOS , FAR_TOT_DESCTO, FAR_DESCTO FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ?  AND FAR_ESTADO <> 'E' ORDER BY FAR_NUMSER, FAR_NUMFAC, FAR_NUMSEC "
Set FARUSU = CN.CreateQuery("", pub_cadena)
FARUSU(0) = 0
FARUSU(1) = 0
FARUSU(2) = 0
FARUSU(3) = 0
FARUSU(4) = 0
Set far_codusu = FARUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
FARUSU(0) = LK_CODCIA
FARUSU(1) = LOC_TIPMOV
FARUSU(2) = d_fecha.Caption
FARUSU(3) = txtSerie.Text
FARUSU(4) = txtNumfac.Text
far_codusu.Requery
fila = 1
Do Until far_codusu.EOF
 far_codusu.Edit
 far_codusu!FAR_DESCTO = val(grid_fac2.TextMatrix(fila, 6))
 far_codusu!FAR_PORDESCTOS = val(grid_fac2.TextMatrix(fila, 6))
 far_codusu!FAR_TOT_DESCTO = cap_valor
 far_codusu.Update
 far_codusu.MoveNext
 fila = fila + 1
Loop

MsgBox "Ok. Descto  Modificado." & Chr(13) & "Todos los articulos del documento deben ser costeados nuevamemte.", 48, Pub_Titulo
txtnumfac_KeyPress 13
 

End Sub

Private Sub d_dire_DblClick()
FRADIRE.Visible = True
Azul txtdire, txtdire
End Sub

Private Sub d_flete_DblClick()
Dim cap_valor
Dim wcanti  As Currency
Dim wpeso  As Currency
If LK_EMP <> "3AA" Then Exit Sub
If LOC_TIPMOV <> 20 Then Exit Sub

 cap_valor = InputBox("Modificación de Flete en Mercaderia  S/. = ", "Fletes en S/.", d_flete.Caption)
 If cap_valor = "" Then Exit Sub
 If val(cap_valor) = 0 Then
  pub_mensaje = "Valor 0.00 para el Flete... desea continuar... "
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta <> vbYes Then
   Exit Sub
  End If
 End If
 
 wcanti = 0
 wpeso = 0
 fila = 1
 Do Until fila = grid_fac2.rows
   wcanti = wcanti + val(grid_fac2.TextMatrix(fila, 2))
   wpeso = wpeso + val(grid_fac2.TextMatrix(fila, 8))
   fila = fila + 1
 Loop
   

pub_mensaje = "Flete x Cantidad (Si), x Peso (No)...."
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
   ww_desc = 0
   fila = 1
   Do Until fila = grid_fac2.rows
     If Pub_Respuesta = vbYes Then
       If wcanti <> 0 Then
           grid_fac2.TextMatrix(fila, 7) = Format((val(cap_valor) / wcanti) * val(grid_fac2.TextMatrix(fila, 2)), "0.00")
       End If
     Else
       If wpeso <> 0 Then
           grid_fac2.TextMatrix(fila, 7) = Format((val(cap_valor) / wpeso) * grid_fac2.TextMatrix(fila, 8), "0.00")
       End If
     End If
     ww_desc = ww_desc + val(grid_fac2.TextMatrix(fila, 7))
     fila = fila + 1
   Loop

pub_mensaje = "Chequear los datos del calculo. Total de Flete : " & Format(ww_desc, "0.00") & " - Confirmar la modificación ? "
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   txtnumfac_KeyPress 13
   Exit Sub
End If
   

'Exit Sub


pub_cadena = "SELECT FAR_TOT_FLETE, FAR_FLETE FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ?  AND FAR_ESTADO <> 'E' ORDER BY FAR_NUMSER, FAR_NUMFAC, FAR_NUMSEC "
Set FARUSU = CN.CreateQuery("", pub_cadena)
FARUSU(0) = 0
FARUSU(1) = 0
FARUSU(2) = 0
FARUSU(3) = 0
FARUSU(4) = 0
Set far_codusu = FARUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
FARUSU(0) = LK_CODCIA
FARUSU(1) = LOC_TIPMOV
FARUSU(2) = d_fecha.Caption
FARUSU(3) = txtSerie.Text
FARUSU(4) = txtNumfac.Text
far_codusu.Requery
fila = 1
Do Until far_codusu.EOF
 far_codusu.Edit
 far_codusu!FAR_FLETE = val(grid_fac2.TextMatrix(fila, 7))
 far_codusu!FAR_TOT_FLETE = ww_desc
 far_codusu.Update
 far_codusu.MoveNext
 fila = fila + 1
Loop

MsgBox "Ok. Fletes  Modificado." & Chr(13) & "Todos los articulos del documento deben ser costeados nuevamemte.", 48, Pub_Titulo
txtnumfac_KeyPress 13
 
End Sub





Private Sub d_usuario_Click()
If Trim(d_usuario.Caption) = "" Then Exit Sub
If Trim(LK_CODUSU) = "ADMIN" Or Trim(LK_CODUSU) = "SUPERVISOR" Then
Else
Exit Sub
End If
Dim FARUSU As rdoQuery
Dim far_codusu As rdoResultset
Dim ALLUSU As rdoQuery
Dim all_codusu As rdoResultset
wcodusu = InputBox("Ingrese Usuario", "Cambio de Usuario", d_usuario.Caption)
If wcodusu = "" Then Exit Sub
wcodusu = UCase(wcodusu)
PSUSU_LLAVE(0) = UCase(wcodusu)
usu_llave.Requery
If usu_llave.EOF Then
   MsgBox "Usuario no Existe", 48, Pub_Titulo
   Exit Sub
End If
  

pub_cadena = "SELECT ALL_CODUSU FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_TIPMOV = ? AND ALL_FECHA_DIA = ? AND ALL_NUMSER = ? AND ALL_NUMFAC = ? ORDER BY ALL_NUMFAC"
Set ALLUSU = CN.CreateQuery("", pub_cadena)
ALLUSU(0) = 0
ALLUSU(1) = 0
ALLUSU(2) = 0
ALLUSU(3) = 0
ALLUSU(4) = 0
Set all_codusu = ALLUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
ALLUSU(0) = LK_CODCIA
ALLUSU(1) = LOC_TIPMOV
ALLUSU(2) = d_fecha.Caption
ALLUSU(3) = txtSerie.Text
ALLUSU(4) = txtNumfac.Text

all_codusu.Requery
Do Until all_codusu.EOF
' Print all_codusu!ALL_FECHA_DIA
' Print all_codusu!ALL_FLAG_SO
 all_codusu.Edit
 all_codusu!all_codusu = wcodusu
 all_codusu.Update
 all_codusu.MoveNext
 Loop
pub_cadena = "SELECT FAR_CODUSU FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ? ORDER BY FAR_NUMFAC"
Set FARUSU = CN.CreateQuery("", pub_cadena)
FARUSU(0) = 0
FARUSU(1) = 0
FARUSU(2) = 0
FARUSU(3) = 0
FARUSU(4) = 0
Set far_codusu = FARUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
FARUSU(0) = LK_CODCIA
FARUSU(1) = LOC_TIPMOV
FARUSU(2) = d_fecha.Caption
FARUSU(3) = txtSerie.Text
FARUSU(4) = txtNumfac.Text

far_codusu.Requery
Do Until far_codusu.EOF
 far_codusu.Edit
 far_codusu!far_codusu = wcodusu
 far_codusu.Update
 far_codusu.MoveNext
Loop

MsgBox "Cambio efectuado.", 48, Pub_Titulo

End Sub

Private Sub FCONT_DblClick()
Dim wfecha
wfecha = InputBox("Modificar Fecha Contable ", "Fecha Contable", "")
If wfecha = "" Then Exit Sub
If Not IsDate(wfecha) Then
 MsgBox "Fecha Incorrecta.., No Procede ", 48, Pub_Titulo
 Exit Sub
End If
pub_cadena = "SELECT ALL_FECHA_PRO  FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_TIPMOV = ? AND ALL_FECHA_DIA = ? AND ALL_NUMSER = ? AND ALL_NUMFAC = ? ORDER BY ALL_NUMFAC"
Set ALLUSU = CN.CreateQuery("", pub_cadena)
ALLUSU(0) = 0
ALLUSU(1) = 0
ALLUSU(2) = 0
ALLUSU(3) = 0
ALLUSU(4) = 0
Set all_codusu = ALLUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
ALLUSU(0) = LK_CODCIA
ALLUSU(1) = LOC_TIPMOV
ALLUSU(2) = d_fecha.Caption
ALLUSU(3) = txtSerie.Text
ALLUSU(4) = txtNumfac.Text
all_codusu.Requery
Do Until all_codusu.EOF
 all_codusu.Edit
 all_codusu!ALL_FECHA_PRO = Format(wfecha, "dd/mm/yyyy")
 all_codusu.Update
 all_codusu.MoveNext
 Loop

pub_cadena = "SELECT FAR_FECHA_PRO FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ? ORDER BY FAR_NUMFAC"
Set FARUSU = CN.CreateQuery("", pub_cadena)
FARUSU(0) = 0
FARUSU(1) = 0
FARUSU(2) = 0
FARUSU(3) = 0
FARUSU(4) = 0
Set far_codusu = FARUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
FARUSU(0) = LK_CODCIA
FARUSU(1) = LOC_TIPMOV
FARUSU(2) = d_fecha.Caption
FARUSU(3) = txtSerie.Text
FARUSU(4) = txtNumfac.Text
far_codusu.Requery
Do Until far_codusu.EOF
 far_codusu.Edit
 far_codusu!FAR_fecha_pro = Format(wfecha, "dd/mm/yyyy")
 far_codusu.Update
 far_codusu.MoveNext
Loop


FCONT.Caption = "Fec. Contable : " & Format(wfecha, "dd/mm/yyyy")

End Sub

Private Sub Form_Load()
Dim SQL As String

LlenadoCbo cmbMotivo, 100

Unload FORMGEN
'Unload FORM_GRIFO
'*********************************************************
SQL = "SELECT DIRCOMP FROM DIRCLI " & _
"WHERE CODCIA=? AND DIRCLI=? AND CODCLI=? AND CP=?"
Set PS = CN.CreateQuery("", SQL)
  PS.rdoParameters(0) = " "
  PS.rdoParameters(1) = 0
  PS.rdoParameters(2) = 0
  PS.rdoParameters(3) = " "
  Set rs = PS.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
'*********************************************************
loc_flag_espera = ""
LOC_ARROZ = ""
pub_cadena = "SELECT * FROM TRANSPORTE WHERE TRN_KEY = ? ORDER BY TRN_KEY"
Set PS_TRA = CN.CreateQuery("", pub_cadena)
PS_TRA(0) = 0
Set llave_trans = PS_TRA.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

LOC_NUMFAC_FIN = 0
pub_cadena = "SELECT * FROM TRANSPORTE WHERE TRN_KEY >= ? ORDER BY TRN_NOMBRE"
Set pstransporte = CN.CreateQuery("", pub_cadena)
pstransporte.rdoParameters(0) = 0
Set TRANSPORTE = pstransporte.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
pstransporte(0) = 0
TRANSPORTE.Requery
TRANS.Clear
Do Until TRANSPORTE.EOF
    TRANS.AddItem Trim(TRANSPORTE!TRN_NOMBRE) & String(80, " ") & TRANSPORTE!TRN_KEY
    TRANSPORTE.MoveNext
Loop
TRANS.Enabled = False
pub_cadena = "SELECT * FROM facart WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_NUMSER = ? AND FAR_FBG=? AND FAR_NUMFAC = ?    ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_NUMSER, FAR_FBG, FAR_NUMFAC, FAR_NUMSEC"
Set PSFAR = CN.CreateQuery("", pub_cadena)
PSFAR.rdoParameters(0) = 0
PSFAR.rdoParameters(1) = 0
PSFAR.rdoParameters(2) = 0
PSFAR.rdoParameters(3) = 0
PSFAR.rdoParameters(4) = 0
Set far_r = PSFAR.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
'SELE_DOCU
pub_cadena = "SELECT FAR_numfac FROM facart WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_FBG = ? AND FAR_NUMSER = ? AND FAR_ESTADO <> 'E' AND FAR_CP = ?  ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_FBG , FAR_NUMSER, FAR_NUMFAC DESC"
Set PSFAR_CONSUL = CN.CreateQuery("", pub_cadena)
PSFAR_CONSUL.rdoParameters(0) = 0
PSFAR_CONSUL.rdoParameters(1) = " "
PSFAR_CONSUL.rdoParameters(2) = " "
PSFAR_CONSUL.rdoParameters(3) = 0
PSFAR_CONSUL.rdoParameters(4) = " "
PSFAR_CONSUL.MaxRows = 1
Set far_consul = PSFAR_CONSUL.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM CARTERA WHERE CAR_CODCIA = ? AND CAR_CP = ? AND CAR_CODCLIE = ?  AND CAR_FBG = ? AND CAR_NUMSER = ? AND CAR_NUMFAC = ? AND CAR_IMPORTE <> 0 ORDER BY CAR_CODCIA, CAR_CODCLIE"
Set PSCAR_CONSUL = CN.CreateQuery("", pub_cadena)
PSCAR_CONSUL.rdoParameters(0) = " "
PSCAR_CONSUL.rdoParameters(1) = " "
PSCAR_CONSUL.rdoParameters(2) = 0
PSCAR_CONSUL.rdoParameters(3) = " "
PSCAR_CONSUL.rdoParameters(4) = 0
PSCAR_CONSUL.rdoParameters(5) = 0
Set car_consul = PSCAR_CONSUL.OpenResultset(rdOpenKeyset, rdConcurValues)
PUB_CODCIA = "00"
LLENADOS TIPMOV, 4
LOC_TIPMOV = 0
temporal = "X"
cmbFBG.Clear
lblflete.Caption = "Flete"
stbEtiqueta.Panels(5).Text = "Flete"
lblNumfac.Caption = "Nº de Doc."
temporal = ""
tempo_serie = ""
WGUIA_RELA = ""
SERIE_GUIA = ""
If LK_FLAG_GRIFO = "A" Then
End If
If LK_FLAG_FACTURACION = "V" Then
 txtvend.Visible = True
 lblvend.Visible = True
Else
 txtvend.Visible = False
 lblvend.Visible = False
End If
lblvend.Caption = "Vend."
If LK_FLAG_GRIFO = "A" Then
  lblvend.Caption = "Isla"
End If

LLENA_ZONA TxtZonaTrabajo, 20
LLENA_ZONA TxtSubZonaTrabajo, 35

End Sub

Public Sub LLENA_CONSULTA()

    '*** VERFICA SI HAY DATOS , O ESTAN CORRECTOS
    'On Error GoTo SALE
    Dim conteo_cantidad As Currency

    Dim conteo_peso     As Currency

    Dim wnumser_c       As String

    Dim wnumfac_c       As String

    Dim WFECHA_CONT     As Date

    Dim WFECHA_COMPRA   As Date

    Dim wwsigno_car     As Integer

    Dim wwdias          As Integer

    Dim DESC_GRIFO      As Currency

    Dim WISLA

    Dim WRES

    Dim ws_serie      As Integer

    Dim WS_CUENTA     As Integer

    Dim suma_subtotal As Currency

    If LOC_TIPMOV = 10 Then NUMERO.Text = 1
    If LOC_TIPMOV = 100 Then NUMERO.Text = 5
    If LOC_TIPMOV = 3 Then NUMERO.Text = 7

    DESC_GRIFO = 0
    WS_FLETE = 0
    ws_serie = 0
    ws_serie = val(txtSerie.Text)
    CmdAnterior.Enabled = False
    'DoEvents
    cmdSiguiente.Enabled = False
    cmdImp.Enabled = False
    'DoEvents
    cherela.Visible = False
    tguia.Visible = False
    chetrans.Visible = False
    TRANS.Visible = False
    l_fecha_compra.Visible = False
    d_fecha_compra.Visible = False
    FECHA_PART.Visible = False
    l_fecha_compra.Visible = True
    d_fecha_compra.Visible = True

    If LOC_TIPMOV = 100 Or LOC_TIPMOV = 3 Then
        l_fecha_compra.Visible = True
        d_fecha_compra.Visible = True
        FECHA_PART.Visible = True
 
        tguia.Visible = True
             
        cherela.Visible = True
        chetrans.Visible = True
        tguia.ToolTipText = "F/B/G = @GUIA y  G.Rem. = @NUMDOC"
        TRANS.Visible = True
    End If

    If LOC_TIPMOV = 10 Then

        l_fecha_compra.Visible = True
        d_fecha_compra.Visible = True
        FECHA_PART.Visible = True
 
        tguia.Visible = True

        cherela.Visible = True
        chetrans.Visible = True
        tguia.ToolTipText = "F/B/G = @GUIA y  G.Rem. = @NUNDOC"
        TRANS.Visible = True
        PSFAR.rdoParameters(0) = LOC_TIPMOV
        PSFAR.rdoParameters(2) = ws_serie
        PSFAR.rdoParameters(3) = Left(cmbFBG.Text, 1)
    ElseIf LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
        PSFAR.rdoParameters(0) = LOC_TIPMOV
        PSFAR.rdoParameters(2) = ws_serie
        PSFAR.rdoParameters(3) = Left(cmbFBG.Text, 1)

        If LOC_TIPMOV = 97 Then  'ALAN DIJO QUE SALIA ERROR ???
            PSFAR.rdoParameters(5) = Right(cmbFBG.Text, 1)
        End If

    ElseIf LOC_TIPMOV = 20 Or LOC_TIPMOV = 99 Then

        If LOC_TIPMOV = 20 And Left(cmbFBG.Text, 1) <> "K" Then
            If Left(cmbFBG.Text, 1) = "F" Then
                PSFAR.rdoParameters(0) = LOC_TIPMOV
                PSFAR.rdoParameters(1) = LK_CODCIA
                PSFAR.rdoParameters(2) = txtSerie.Text
                PSFAR.rdoParameters(3) = val(txtNumfac.Text)
            Else
                PSFAR.rdoParameters(0) = LOC_TIPMOV
                PSFAR.rdoParameters(1) = LK_CODCIA
                PSFAR.rdoParameters(2) = txtNumfac.Text
            End If

        Else
            PSFAR.rdoParameters(0) = LOC_TIPMOV
            PSFAR.rdoParameters(2) = ws_serie

            If LOC_TIPMOV = 99 Then
                PSFAR.rdoParameters(3) = Left(cmbFBG.Text, 1)
            Else
                PSFAR.rdoParameters(3) = " "
            End If
        End If

    Else
        PSFAR.rdoParameters(0) = LOC_TIPMOV
        PSFAR.rdoParameters(2) = ws_serie
        PSFAR.rdoParameters(3) = ""
    End If

    If LOC_TIPMOV = 20 And Left(cmbFBG.Text, 1) <> "K" Then
    Else
        PSFAR.rdoParameters(1) = LK_CODCIA
        PSFAR.rdoParameters(4) = val(txtNumfac.Text)
    End If

    far_r.Requery

    If far_r.EOF Then
        wflag_docu = "A"
        CmdAnterior.Enabled = True
        cmdSiguiente.Enabled = True
        cmdImp.Enabled = True

        Exit Sub

    Else

        If LK_FLAG_SOS = "A" And far_r!FAR_FLAG_SO <> "A" Then
            wflag_docu = "A"
            CmdAnterior.Enabled = True
            cmdSiguiente.Enabled = True
            cmdImp.Enabled = True

            Exit Sub

        End If
    End If

    tempo_serie = Trim(txtSerie.Text)
    PB.Value = 0
    PB.Min = 0
    PB.max = far_r.RowCount + 3
    PB.Visible = True
    'DoEvents
    LIMPIA_DOCU
    'MsgBox far_r!far_NUMOPER
    'MsgBox far_r!far_CONCEPTO
    wflag_activo = 0
    WS_CUENTA = far_r.RowCount
    fila = 0

    Do Until far_r.EOF

        If Nulo_Valors(far_r!far_estado) <> "E" Then
            wflag_activo = 1
            GoTo muestra

            Exit Do

        End If

        far_r.MoveNext
    Loop

    far_r.MoveLast
    fila = far_r!FAR_NUMSEC

    Do Until far_r.BOF
        wfecha_anu = far_r!FAR_fecha

        If fila <> far_r!FAR_NUMSEC Then
            far_r.MoveNext
            GoTo muestra
        End If

        fila = fila - 1
        far_r.MovePrevious
    Loop

    far_r.MoveFirst

muestra:

    'd_fecha.Caption = Format(far_r!FAR_fecha, "dd/mm/yyyy")
    'If Nulo_Valors(far_r!far_estado) = "E" Then
    '  If wfecha_anu <> LK_FECHA_DIA Then
    '    LBLEXTORNO.Caption = "DOCUMENTO A N U L A D O"
    '    d_fecha.Caption = Format(wfecha_anu, "dd/mm/yyyy")
    '  Else
    '    If far_r!FAR_fecha = LK_FECHA_DIA Then
    '      LBLEXTORNO.Caption = "DOCUMENTO EXTORNADO"
    '    End If
    '  End If
    '  LBLEXTORNO.Visible = True
    'Else
    '  LBLEXTORNO.Visible = False
    'End If
    DOC102 = ""

    SQ_OPER = 1
    pu_codcia = LK_CODCIA

    If LOC_TIPMOV = 93 Or LOC_TIPMOV = 5 Or LOC_TIPMOV = 6 Or LOC_TIPMOV = 100 Or LOC_TIPMOV = 101 Then
        lbldomicilio.Caption = "Destino :"
        lbldireccion.Caption = "Concepto :"
        'QUITADO 30/11/2001
        d_dire.Caption = Trim(far_r!far_concepto)
        d_domicilio.Caption = Trim(far_r!far_subtra)

        If LOC_TIPMOV = 5 Then
            tguia.Visible = True
            cherela.Visible = True
            chetrans.Visible = True
            tguia.ToolTipText = "F/B/G = @GUIA y  G.Rem. = @NUNDOC"
            TRANS.Visible = True
        End If

    Else
        lbldireccion.Caption = "Dirección Entrega:"
        lbldomicilio.Caption = "Domicilio:"
 
    End If

    If LOC_TIPMOV = 102 Or OC_TIPMOV = 100 Or LOC_TIPMOV = 10 Or LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
  
        txtdocu.Visible = True
        pu_cp = "C"
        lbldomicilio.Caption = "Domicilio:"
        PU_FBG = far_r!far_fbg
        WGUIA_RELA = far_r!far_numguia
        SERIE_GUIA = far_r!far_serguia

        If LK_EMP <> "HER" Then
  
            txtdocu.Caption = "G/ " & far_r!far_serguia & " - " & far_r!far_numguia & " - Nº.O/C :" & Trim(Nulo_Valors(far_r!FAR_OC))
        Else
            txtdocu.Caption = "G/ " & far_r!far_serguia & " - " & far_r!far_numguia
        End If

        If val(WGUIA_RELA) <> 0 Then
            cherela.Value = 1
        End If

        If LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
            txtdocu.Caption = Trim(far_r!far_concepto)  ' serguia & " - " & far_r!far_numguia
            'txtdocu.Caption = all_llave!all_concepto
            pu_cp = Right(cmbFBG.Text, 1)
        End If

        lblsaldo.Visible = True
        d_saldo.Visible = True
        d_moneda.Visible = True
    ElseIf LOC_TIPMOV = 20 Or LOC_TIPMOV = 99 Or LOC_TIPMOV = 3 Then
        l_fecha_compra.Visible = True
        d_fecha_compra.Visible = True
        lblsaldo.Visible = True
        d_saldo.Visible = True
 
        d_moneda.Visible = True
        'lblsaldo.Visible = False
        'd_saldo.Visible = False
        txtdocu.Visible = True
        lblfac.Visible = True

        If val(far_r!FAR_NUMFAC_C) = 0 Then
            'txtdocu.Caption = " "
            txtdocu.Caption = "G/ " & far_r!far_serguia & " - " & far_r!far_numguia
        Else
            txtdocu.Caption = "F/ " & far_r!FAR_NUMSER_C & " - " & far_r!FAR_NUMFAC_C
        End If

        If LOC_TIPMOV = 99 Then
            PUB_CODCIA = "00"
            PUB_TIPREG = 50
            PUB_NUMTAB = far_r!far_cod_sunat
            LEER_TAB_LLAVE

            If Not tab_llave.EOF Then
                If far_r!FAR_NUMFAC_C <> 0 Then
                    txtdocu.Caption = Trim(str(far_r!far_cod_sunat)) & "-" & Trim(Left(tab_llave!tab_NOMLARGO, 20)) & " / " & far_r!FAR_NUMSER_C & " - " & far_r!FAR_NUMFAC_C
                Else
                    txtdocu.Caption = Trim(tab_llave!tab_NOMLARGO) & " / " & far_r!far_numguia
                End If
            End If
        End If
        
        pu_cp = far_r!FAR_cp
        'pu_cp = "C"
        'If Right(Trim(cmbFBG.Text), 1) = "P" Then
        '  pu_cp = "C"
        'End If
 
        PU_FBG = " "
    Else
        GoTo PASACLI
    End If

    pu_codclie = far_r!far_codclie
    LEER_CLI_LLAVE

    If cli_llave.EOF Then
        MsgBox "Registro de Cliente not :" & far_r!far_codclie, 48, Pub_Titulo
        GoTo SALE
    End If

    PB.Value = PB.Value + 1
    d_Codclie.Caption = Trim(cli_llave!cli_codclie)

   ' If Trim(cli_llave!cli_codclie) = 109 Then
        'd_nomclie.Caption = Trim(far_r!FAR_CLIENTE)    ' GTS p Dirome cambio Trim(far_r!FAR_CLIENTE) por (cli_llave!CLI_NOMBRE)
   ' Else
        d_nomclie.Caption = Trim(cli_llave!CLI_NOMBRE)
   ' End If

    If LOC_TIPMOV = 100 Then
        d_domicilio.Caption = Trim(far_r!far_subtra)
    End If

    SQ_OPER = 1
    PUB_CODCIA = "00"
    'PUB_NUMTAB = cli_llave!CLI_LUGAR_TRAB
    PUB_TIPREG = 25
    LEER_TAB_LLAVE
    WLUGAR = ""

    If Not tab_llave.EOF Then
        WLUGAR = Trim(tab_llave!tab_NOMLARGO)
    End If

    'PUB_NUMTAB = cli_llave!CLI_LUGAR_CASA
    LEER_TAB_LLAVE
    WLUGAR1 = ""

    If Not tab_llave.EOF Then
        WLUGAR1 = Trim(tab_llave!tab_NOMLARGO)
    End If

    PUB_NUMTAB = Nulo_Valor0(cli_llave!cli_TRAB_ZONA)
    PUB_TIPREG = 20
    LEER_TAB_LLAVE
    WZONA = ""

    If Not tab_llave.EOF Then
        WZONA = Trim(tab_llave!tab_NOMLARGO)
    End If

    PUB_NUMTAB = Nulo_Valor0(cli_llave!CLI_CASA_ZONA)
    LEER_TAB_LLAVE
    WZONA1 = ""

    If Not tab_llave.EOF Then
        WZONA1 = Trim(tab_llave!tab_NOMLARGO)
    End If

    'PUB_NUMTAB = cli_llave!cli_TRAB_SUBZONA
    PUB_TIPREG = 35
    LEER_TAB_LLAVE
    WSUBZONA = ""

    If Not tab_llave.EOF Then
        WSUBZONA = Trim(tab_llave!tab_NOMLARGO)
    End If

    PUB_NUMTAB = cli_llave!CLI_ZONA_NEW
    LEER_TAB_LLAVE
    WSUBZONA1 = ""

    If Not tab_llave.EOF Then
        WSUBZONA1 = Trim(tab_llave!tab_NOMLARGO)
    End If

    'QUITADO 30/11/2001
    'd_dire.Caption = Trim(WLUGAR) + " " + Trim(cli_llave!CLI_TRAB_DIREC) + " # " + Trim(cli_llave!CLI_TRAB_NUM) & "  " & WZONA & "  " & WSUBZONA
    'txtdire.Text = Trim(cli_llave!CLI_TRAB_DIREC)
    'txtnum.Text = Trim(cli_llave!CLI_TRAB_NUM)
    'ASIGNA_INT TxtZonaTrabajo, cli_llave!cli_TRAB_ZONA
    'ASIGNA_INT TxtSubZonaTrabajo, cli_llave!cli_TRAB_SUBZONA
    If LOC_TIPMOV = 100 Then
        d_domicilio.Caption = Trim(far_r!far_subtra)
    Else
        d_domicilio.Caption = Trim(WLUGAR1) + " " + Trim(cli_llave!CLI_CASA_DIREC) + " # " + Trim(cli_llave!CLI_CASA_NUM) & "  " & WZONA1 & "  " & WSUBZONA1
    End If

    If Left(cmbFBG.Text, 1) = "F" Then
        d_ruc.Caption = Trim(Nulo_Valors(cli_llave!cli_ruc_esposo))
        LBLRUC.Visible = True
    Else
        LBLRUC.Visible = False
        d_ruc.Caption = ""
    End If

PASACLI:

    If LK_FLAG_GRIFO = "A" And far_r!FAR_TIPMOV = 20 Then
        d_nomven.Caption = "Turno: " & Format(far_r!far_turno, "00") & Chr(13) & "Nº.Carga: " & far_r!FAR_PEDFAC
    End If

    If far_r!FAR_cp <> "P" Then
        If LOC_TIPMOV = 102 Or LOC_TIPMOV = 10 Or LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
            If LK_FLAG_GRIFO <> "A" Then
                If txtvend.Visible Then
                    txtvend.Text = Nulo_Valor0(far_r!FAR_CODVEN)
                End If

                SQ_OPER = 1
                pu_codcia = LK_CODCIA
                PUB_CODVEN = Nulo_Valor0(far_r!FAR_CODVEN)
                LEER_VEN_LLAVE

                If ven_llave.EOF Then
                    ' MsgBox "Verificar Vendedor.", 48, Pub_Titulo
                Else
                    d_codven.Caption = Nulo_Valor0(Trim(far_r!FAR_CODVEN))
                    d_nomven.Caption = Trim(ven_llave!VEM_NOMBRE)
                End If

            Else

                If txtvend.Visible Then
                    txtvend.Text = far_r!far_ISLA
                End If

                SQ_OPER = 1
                PUB_CODCIA = xCODCIA
                PUB_CODVEN = Nulo_Valor0(far_r!far_ISLA)
                LEER_VEN_LLAVE
                WISLA = " "

                If Not ven_llave.EOF Then
                    WISLA = ven_llave!VEM_NOMBRE
                End If

                SQ_OPER = 1
                PUB_TIPREG = 2103
                PUB_CODCIA = LK_CODCIA
                PUB_NUMTAB = val(far_r!FAR_CODVEN)
                LEER_TAB_LLAVE
                WRES = " "

                If Not tab_llave.EOF Then
                    WRES = Trim(tab_llave!tab_NOMLARGO)
                End If

                d_codven.Caption = "-"
                d_nomven.Caption = WISLA & " / " & WRES & " - " & "Turno: " & Format(far_r!far_turno, "00")
            End If
        End If
    End If

    d_usuario.Caption = Trim(Nulo_Valors(far_r!far_codusu))
    pu_cp = far_r!FAR_cp
    pu_codcia = LK_CODCIA
    PU_NUMSER = far_r!far_numser
    PU_NUMFAC = far_r!far_numfac
    LEER_CAR_CONSUL

    If Not car_consul.EOF Then
        d_saldo.Caption = Format(car_consul!car_importe, "##,##0.000")
        d_newvcto.Caption = Format(car_consul!car_fecha_vcto, "dd/mm/yyyy")
    End If

    If LOC_TIPMOV = 102 Then
        SQ_OPER = 8
        PU_TIPMOV = 10
        pu_codcia = LK_CODCIA
        PU_FBG = "F"
        PU_FBG2 = "B"
        LEER_FAR_LLAVE

        If far_menor4.EOF Then

            DOC102 = ""
        Else

            DOC102 = far_menor4!far_fbg & "/ " & Format(far_menor4!far_numser, "000") & "-" & far_menor4!far_numfac
        End If

        txtdocu.Visible = True
        txtdocu.Caption = " Doc. Relacionado. = " & DOC102
    End If

    PB.Value = PB.Value + 1
    wwsigno_car = Nulo_Valor0(far_r!far_signo_car)
    wwdias = Nulo_Valor0(far_r!FAR_DIAS)
    FECHA_PART.Text = Format(far_r!FAR_fecha_compra, "dd/mm/yyyy")

    ' configurar despues acv
    If wwsigno_car = 1 And wwdias <> 0 Then
        If LOC_TIPMOV = 10 Then
            d_condicion.Caption = "VENTA AL CREDITO"
        ElseIf LOC_TIPMOV = 20 Or LOC_TIPMOV = 99 Then
            d_condicion.Caption = far_r!far_subtra ' far_r!FAR_secuencia ' "COMPRA AL CREDITO"
            d_fecha_compra.Caption = Format(far_r!FAR_fecha_compra, "dd/mm/yyyy")
        Else
            d_condicion.Caption = Left(TIPMOV.Text, 40)
        End If

        d_dias.Caption = val(far_r!FAR_DIAS)
        d_fechaV.Caption = Format(DateAdd("d", val(far_r!FAR_DIAS), far_r!FAR_fecha_compra), "dd/mm/yyyy")
    Else

        If LOC_TIPMOV = 10 Then
            d_condicion.Caption = "VENTA AL CONTADO"
            frmdocu.d_fechaV.Caption = frmdocu.d_fecha.Caption
     
        ElseIf LOC_TIPMOV = 20 Then
            d_condicion.Caption = far_r!far_subtra ' "COMPRA AL CONTADO"
        Else
            d_condicion.Caption = Left(TIPMOV.Text, 40)

            If LOC_TIPMOV = 99 Or LOC_TIPMOV = 5 Or LOC_TIPMOV = 6 Then
                d_condicion.Caption = Nulo_Valors(far_r!far_subtra)
            End If
        End If
    End If

    If LK_FLAG_GRIFO = "A" And LOC_TIPMOV = 10 Then
        If val(Nulo_Valor0(far_r!far_signo_car)) <> 1 Then
            d_condicion.Caption = "VENTA AL CONTADO"
        Else
            d_condicion.Caption = "VENTA AL CREDITO"
        End If
    End If

    'lblven.Caption = "Vendedor"
    If Nulo_Valors(far_r!FAR_MONEDA) = "D" Then
        d_moneda.Caption = "$."
        tmoneda.Text = "US$."

        '  lblven.Visible = True
        If LOC_TIPMOV = 20 Then
            '   lblven.Caption = "T. Cambio:"
            d_codven.Caption = val(Nulo_Valor0(far_r!FAR_tipo_cambio))
        End If

    Else
        d_moneda.Caption = "S/."
        tmoneda.Text = "S/."
    End If

    d_fecha_compra.Caption = Format(far_r!FAR_fecha_compra, "dd/mm/yyyy")
    d_fecha.Caption = Format(far_r!FAR_fecha, "dd/mm/yyyy")
    wfecha_anu = far_r!FAR_fecha

    If Nulo_Valors(far_r!far_estado) = "E" Then
        If wfecha_anu <> LK_FECHA_DIA Then
            LBLEXTORNO.Caption = "DOCUMENTO A N U L A D O"
            d_fecha.Caption = Format(wfecha_anu, "dd/mm/yyyy")
        Else

            If far_r!FAR_fecha = LK_FECHA_DIA Then
                LBLEXTORNO.Caption = "DOCUMENTO EXTORNADO"
            End If
        End If

        LBLEXTORNO.Visible = True
    Else
        LBLEXTORNO.Visible = False
    End If

    cabe_grid
    PB.Value = PB.Value + 1
    fila = 0
    WS_BRUTO = 0
    SUB_CANT = 0
    subtotal = 0
    PUB_DESCTO = 0
    grid_fac2.rows = 1
    fila = 0
    suma_subtotal = 0
    LOC_ARROZ = ""
    '**********************
    'comienza llenado

    PS(0) = LK_CODCIA
    PS(1) = far_r!far_key_dircli
    PS(2) = far_r!far_codclie
    PS(3) = far_r!FAR_cp
    rs.Requery

    If rs.EOF Then
        'Exit Sub
    End If

    Do Until rs.EOF
        d_dire = rs!dircomp
        rs.MoveNext
    Loop

    'termina llenar direccción
 
    conteo_cantidad = 0
    conteo_peso = 0

    Do Until far_r.EOF

        If LOC_TIPMOV = 20 And val(far_r!far_signo_arm) = -1 And far_r!far_estado = "E" Then GoTo NADA
        If (LOC_TIPMOV = 97 Or LOC_TIPMOV = 10) And val(far_r!far_codart) = 0 Then
            grid_fac2.rows = grid_fac2.rows + 3
            fila = fila + 2
            SQ_OPER = 1
            PUB_FECHA = far_r!FAR_fecha
            pu_codcia = LK_CODCIA
            LEER_ALL_LLAVE

            Do Until all_llave.EOF

                If far_r!FAR_NUMOPER = all_llave!ALL_NUMOPER Then Exit Do
                all_llave.MoveNext
            Loop
     
            '     If Not all_llave.EOF Then grid_fac2.TextMatrix(fila, 0) = all_llave!ALL_CONCEPTO
            '     fila = fila + 1
            grid_fac2.TextMatrix(fila, 0) = all_llave!all_concepto
            grid_fac2.ColWidth(0) = 6500
            GoTo pasa
        End If

        If LOC_TIPMOV = 98 And val(far_r!far_codart) = 0 Then
            grid_fac2.rows = grid_fac2.rows + 2
            fila = fila + 2
            grid_fac2.TextMatrix(fila, 0) = far_r!far_concepto
            grid_fac2.ColWidth(0) = 6500
            GoTo pasa
        End If

        If LOC_TIPMOV = 99 And val(far_r!far_codart) = 0 Then
            grid_fac2.rows = grid_fac2.rows + 1
            grid_fac2.TextMatrix(grid_fac2.rows - 1, 0) = "Concepto: " & Trim(far_r!far_concepto)
            grid_fac2.ColWidth(0) = 6500
            SQ_OPER = 1
            pu_codcia = LK_CODCIA
            PUB_FECHA = far_r!FAR_fecha
            LEER_ALL_LLAVE

            Do Until all_llave.EOF

                If all_llave!ALL_NUMOPER = far_r!FAR_NUMOPER Then

                    Exit Do

                End If

                all_llave.MoveNext
            Loop

            If all_llave.EOF Then
                MsgBox "Verificar...DATOS ", 48, Pub_Titulo

                Exit Sub

            End If

            If all_llave!ALL_SIGNO_CCM <> 0 Then
                SQ_OPER = 1
                PUB_CODBAN = all_llave!all_codban
                pu_codcia = LK_CODCIA
                LEER_CCM_LLAVE
                grid_fac2.rows = grid_fac2.rows + 1
                grid_fac2.TextMatrix(grid_fac2.rows - 1, 0) = "Banco: " & Trim(ccm_llave!CCM_NOMBRE) & " Ch/. " & all_llave!ALL_CHESER & " " & all_llave!all_chenum
            End If

            ''If all_llave!ALL_SIGNO_CCM <> 0 Then
            grid_fac2.rows = grid_fac2.rows + 1
            grid_fac2.TextMatrix(grid_fac2.rows - 1, 0) = "Relación Contable: "
            grid_fac2.rows = grid_fac2.rows + 1
            grid_fac2.TextMatrix(grid_fac2.rows - 1, 0) = "'" & all_llave!ALL_CTAG1 & " = " & Format(all_llave!ALL_IMPG1, "#,#00.00")
            grid_fac2.rows = grid_fac2.rows + 1

            If all_llave!ALL_IMPG2 <> 0 Then
                grid_fac2.TextMatrix(grid_fac2.rows - 1, 0) = "'" & all_llave!ALL_CTAG2 & " = " & Format(all_llave!ALL_IMPG2, "#,#00.00")
            End If

            grid_fac2.rows = grid_fac2.rows + 1
            grid_fac2.TextMatrix(grid_fac2.rows - 1, 0) = "' Fec. Cancelación: " & all_llave!ALL_FECHA_CAN
      
            grid_fac2.rows = grid_fac2.rows + 1
            grid_fac2.TextMatrix(grid_fac2.rows - 1, 0) = "' Fec.Contable: " & all_llave!ALL_FECHA_PRO

            ''End If
      
            GoTo pasa
        End If

        PB.Value = PB.Value + 1
        grid_fac2.rows = grid_fac2.rows + 1
        fila = fila + 1
        PUB_KEY = far_r!far_codart
        pu_codcia = LK_CODCIA
        SQ_OPER = 1
        LEER_ART_LLAVE

        If art_LLAVE.EOF Then
            MsgBox "Error Grave en arti..."
        End If

        grid_fac2.TextMatrix(fila, 0) = Trim(art_LLAVE!art_nombre)
        grid_fac2.TextMatrix(fila, 1) = Left(Trim(art_LLAVE!ART_ALTERNO), 8)
        grid_fac2.TextMatrix(fila, 2) = far_r!far_cantidad / far_r!FAR_equiv
        grid_fac2.TextMatrix(fila, 3) = far_r!far_descri
        grid_fac2.TextMatrix(fila, 4) = far_r!FAR_PRECIO
        subtotal = Format(far_r!FAR_PRECIO * (far_r!far_cantidad / far_r!FAR_equiv), "0.00")
        conteo_cantidad = conteo_cantidad + Format((far_r!far_cantidad / far_r!FAR_equiv), "0.00")
        conteo_peso = conteo_peso + Format(((far_r!far_PESO) * (far_r!far_cantidad) / far_r!FAR_equiv), "0.00")

        If LK_FLAG_GRIFO = "A" Then
            grid_fac2.TextMatrix(fila, 6) = Format(far_r!FAR_DESCTO, "0.00")
            DESC_GRIFO = DESC_GRIFO + far_r!FAR_DESCTO
        Else

            If LOC_TIPMOV = 20 Then
                grid_fac2.TextMatrix(fila, 6) = Trim(far_r!FAR_PORDESCTOS)
            Else
                ' grid_fac2.TextMatrix(fila, 6) = Left(Trim(far_r!FAR_PORCDESCSUCE), 120) '& "%" ' Format(far_r!FAR_PORDESCTO1, "0.00") & "%"
            End If
        End If

        If LOC_TIPMOV = 102 Then
            grid_fac2.TextMatrix(fila, 6) = far_r!far_signo_arm
        End If

       ' grid_fac2.TextMatrix(fila, 7) = Trim(art_LLAVE!art_cuenta_contab)  'gts aca lleno color para T Reyes
        'grid_fac2.TextMatrix(fila, 8) = Trim(art_LLAVE!art_cuenta_contab_c)  'gts aca lleno marca para T Reyes
        subtotal = redondea(subtotal)
        suma_subtotal = suma_subtotal + subtotal
        grid_fac2.TextMatrix(fila, 5) = subtotal
        SUB_CANT = SUB_CANT + (far_r!far_cantidad / far_r!FAR_equiv)
pasa:
        WS_BRUTO = far_r!FAR_BRUTO
        WS_DESCTO = far_r!FAR_TOT_DESCTO
        WS_IMPTO = far_r!far_IMPTO
        WS_GASTOS = far_r!FAR_TOT_FLETE
        '  If LK_EMP = "HER" And far_r!FAR_TOT_FLETE <> 0 Then
        WS_FLETE = Nulo_Valor0(far_r!FAR_ICBPER)
        WFECHA_COMPRA = far_r!FAR_fecha_compra
        wnumfac_c = Nulo_Valor0(far_r!FAR_NUMFAC_C)

        If LK_EMP <> "HER" And Not IsNull(far_r!FAR_fecha_pro) Then
            WFECHA_CONT = far_r!FAR_fecha_pro
        End If
   
        If far_r!FAR_EX_IGV = "E" Then LOC_ARROZ = "A"
        ' End If
NADA:
        far_r.MoveNext
    Loop

    d_fecha_compra.Caption = Format(WFECHA_COMPRA, "dd/mm/yyyy")
    FCONT.Caption = "Fec. Contable : " & Format(WFECHA_COMPRA, "dd/mm/yy")  'CAMBIO GTS P DIROME WFECHA_CONT POR WFECHA_COMPRA

    If val(wnumfac_c) <> 0 And LOC_TIPMOV = 10 Then
        txtdocu.Caption = txtdocu.Caption + " / Cambio Doc. = " & wnumser_c & " - " & wnumfac_c
    End If

    If Trim(LBLEXTORNO.Caption) = "" Then
        d_subtotal.Caption = WS_BRUTO '- WS_IMPTO + WS_DESCTO
        d_descto.Caption = WS_DESCTO

        If LK_FLAG_GRIFO = "A" Then d_descto.Caption = Format(DESC_GRIFO, "0.00")
        d_flete.Caption = WS_FLETE
        d_impto.Caption = WS_IMPTO
        d_gastos.Caption = WS_GASTOS

        If WS_BRUTO = WS_DESCTO + WS_GASTOS Then
            d_neto.Caption = "0.000"
        Else

            If LK_EMP = "3AA" Then
                d_neto.Caption = Format(WS_BRUTO + WS_IMPTO, "0.000")
            Else

                If LOC_TIPMOV = 10 Then
                   ' d_neto.Caption = Format(WS_FLETE + WS_BRUTO + WS_IMPTO - WS_GASTOS - WS_DESCTO, "0.000") 'cambie mic (+) WS_GASTOS
                    d_neto.Caption = Format(WS_BRUTO + WS_IMPTO + WS_FLETE - WS_DESCTO + WS_GASTOS, "0.000") 'cambie mic (+) WS_GASTOS
                Else
                    d_neto.Caption = Format(WS_BRUTO + WS_IMPTO - WS_DESCTO + WS_GASTOS, "0.000")
                End If
            End If
        End If

        'If LK_EMP = "HER" And LOC_TIPMOV = 10 Then d_neto.Caption = Format(suma_subtotal, "0.00") ***MIC
        If LOC_TIPMOV = 93 Then
            d_neto.Caption = Format(suma_subtotal, "0.000")
        End If
    End If

    CmdAnterior.Enabled = True
    cmdSiguiente.Enabled = True
    cmdImp.Enabled = True
    PB.Visible = False

    If cherela.Visible And LK_EMP <> "HER" Then cherela_Click
    LOC_NUMFAC_FIN = val(txtNumfac.Text)

    If LK_CODUSU = "ADMIN" And (LOC_TIPMOV = 6 Or LOC_TIPMOV = 5) Then
        ' MsgBox "CAntidad :  " & conteo_cantidad, 48, Pub_Titulo
    End If

    If LK_CODUSU = "PEDRO" And LOC_TIPMOV = 10 And LK_CODCIA = "01" Then  'agregado GTS para Eurotubo
        MsgBox "Peso Total Kgs. :  " & conteo_peso, 48, Pub_Titulo
    End If
   
    Exit Sub

SALE:
    MsgBox Err.Description, 48, Pub_Titulo

    Resume Next

    LIMPIA_DOCU
End Sub
Public Sub cabe_grid()
If LOC_TIPMOV = 96 Then
   grid_fac2.Clear
   grid_fac2.Cols = 15
   grid_fac2.TextMatrix(0, 0) = "Cliente"
   grid_fac2.TextMatrix(0, 1) = "Codigo"
   grid_fac2.TextMatrix(0, 2) = "Vend"
   grid_fac2.TextMatrix(0, 3) = "F/LET"
   grid_fac2.TextMatrix(0, 4) = "F-B-G"
   grid_fac2.TextMatrix(0, 5) = "Serie"
   grid_fac2.TextMatrix(0, 6) = "N.Docum."
   grid_fac2.TextMatrix(0, 7) = "Guia   "
   grid_fac2.TextMatrix(0, 8) = "Saldo    "
   grid_fac2.TextMatrix(0, 9) = "Efectivo"
   grid_fac2.TextMatrix(0, 10) = "Nueva Fecha"
   grid_fac2.TextMatrix(0, 11) = "Nº. Cheque"
   grid_fac2.TextMatrix(0, 12) = "Importe"
   grid_fac2.TextMatrix(0, 13) = "Banco"
   grid_fac2.TextMatrix(0, 14) = "Fec.Cobrar"
   
   grid_fac2.ColWidth(0) = 2000
   grid_fac2.ColWidth(1) = 700
   grid_fac2.ColWidth(2) = 400
   grid_fac2.ColWidth(3) = 500
   grid_fac2.ColWidth(4) = 400
   grid_fac2.ColWidth(5) = 500
   grid_fac2.ColWidth(6) = 800
   grid_fac2.ColWidth(7) = 1
   grid_fac2.ColWidth(8) = 1100
   grid_fac2.ColWidth(9) = 1100
   grid_fac2.ColWidth(10) = 1200
   grid_fac2.ColWidth(11) = 1200
   grid_fac2.ColWidth(12) = 1200
   grid_fac2.ColWidth(13) = 1200
 
Else
   grid_fac2.Clear
   grid_fac2.Cols = 10
   grid_fac2.TextMatrix(0, 0) = "Descripción"
   grid_fac2.TextMatrix(0, 1) = "Codigo"
   grid_fac2.TextMatrix(0, 2) = "Cantidad"
   grid_fac2.TextMatrix(0, 3) = "Unidad"
   grid_fac2.TextMatrix(0, 4) = "Precio"
   grid_fac2.TextMatrix(0, 5) = "Subtotal"
   grid_fac2.TextMatrix(0, 6) = "IMEI"
   grid_fac2.TextMatrix(0, 7) = "Color"
   grid_fac2.TextMatrix(0, 8) = "Marca"
   
   grid_fac2.RowHeight(0) = 385
   grid_fac2.ColWidth(0) = 2500
   grid_fac2.ColWidth(1) = 900
   grid_fac2.ColWidth(2) = 800
   grid_fac2.ColWidth(3) = 900
   grid_fac2.ColWidth(4) = 800
   grid_fac2.ColWidth(5) = 900
   grid_fac2.ColWidth(6) = 2450
   grid_fac2.ColWidth(7) = 0
   grid_fac2.ColWidth(8) = 1200
   grid_fac2.ColWidth(9) = 0
   If LOC_TIPMOV = 20 Then
    grid_fac2.ColWidth(9) = 400
    grid_fac2.TextMatrix(0, 9) = "Flag" '  flag de descto
   End If
   
End If
End Sub



Private Sub grid_fac2_DblClick()
If LOC_TIPMOV <> 20 Then Exit Sub

If grid_fac2.COL = 9 Then
 If Trim(grid_fac2.TextMatrix(grid_fac2.Row, 9)) = "NOT" Then
   grid_fac2.TextMatrix(grid_fac2.Row, 9) = ""
 Else
   grid_fac2.TextMatrix(grid_fac2.Row, 9) = "NOT"
 End If
End If


End Sub

Private Sub TIPMOV_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If txtvend.Visible Then
   txtvend.SetFocus
  Else
   cmbFBG.SetFocus
 End If
End If
End Sub

Private Sub TIPMOV_LostFocus()

cmbFBG.Clear
If Trim(TIPMOV.Text) = "" Then
 LOC_TIPMOV = 0
Else
 LOC_TIPMOV = val(Trim(Right(TIPMOV.Text, 4)))
 cmdImp.Enabled = False
 lblpersona.Visible = True
 d_Codclie.Visible = True
 LBLRUC.Visible = True
 'lblven.Visible = True
 d_codven.Visible = True
 d_condicion.Visible = True
 lblcondicion.Visible = True
 lbldireccion.Visible = True
 If LOC_TIPMOV = 101 Or LOC_TIPMOV = 93 Or LOC_TIPMOV = 20 Or LOC_TIPMOV = 5 Or LOC_TIPMOV = 6 Or LOC_TIPMOV = 10 Or LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
   cmdImp.Enabled = True
 ElseIf LOC_TIPMOV = 96 Then
   cmdImp.Enabled = True
   lblpersona.Visible = False
   d_Codclie.Visible = False
   LBLRUC.Visible = False
'   lblven.Visible = False
   d_codven.Visible = False
   d_condicion.Visible = False
   lblcondicion.Visible = False
   lbldireccion.Visible = False
 End If
 If temporal = "X" Then
  Exit Sub
 End If
 cmbFBG.Clear
 If LOC_TIPMOV = 10 Or LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
  lblflete.Caption = "ICBPER"
  stbEtiqueta.Panels(5).Text = "ICBPER"
  stbEtiqueta.Panels(3).Text = "Servicios"
  lblNumfac.Caption = "Nº de Doc."
  lblpersona.Caption = "Cliente :"
'  lblven.Visible = True
  If LOC_TIPMOV = 97 Then
    cmbFBG.AddItem "N.C. Clientes              C"
    cmbFBG.AddItem "C.N.C. Proveedor             P"
  ElseIf LOC_TIPMOV = 98 Then
    cmbFBG.AddItem "Debito Clientes            C"
    cmbFBG.AddItem "A.Debito Proveedor           P"
  Else
   cmbFBG.AddItem "F = Facturas"
   cmbFBG.AddItem "B = Boletas"
   cmbFBG.AddItem "T = Tickets"
   If LK_FLAG_GRIFO <> "A" Then
    'cmbFBG.AddItem "G = Guias"
    cmbFBG.AddItem "P = Pedidos"
   Else
    cmbFBG.AddItem "P = Auto.C."
    cmbFBG.AddItem "G = O/D.Cont."
    cmbFBG.AddItem "C = O/D.Cred"
   End If
  End If
 ElseIf LOC_TIPMOV = 20 Then
'  lblven.Visible = False
  lblpersona.Caption = "Proveedor :"
  lblNumfac.Caption = "Nºde Kardex"
  lblflete.Caption = "Flete"
  stbEtiqueta.Panels(5).Text = "Flete"
  cmbFBG.AddItem "K = Kardex"
  cmbFBG.AddItem "F = Facturas"
  cmbFBG.AddItem "G = Guias"
 ElseIf LOC_TIPMOV = 99 Or LOC_TIPMOV = 30 Then
  lblpersona.Caption = "Proveedor :"
  lblNumfac.Caption = "Nºde Kardex"
  lblflete.Caption = "Flete"
  stbEtiqueta.Panels(5).Text = "Flete"
  cmbFBG.AddItem "K = Kardex"
  If LOC_TIPMOV = 30 Then cmdImp.Enabled = True
 ElseIf LOC_TIPMOV = 96 Then
  lblNumfac.Caption = "Panilla"
  lblflete.Caption = ""
  stbEtiqueta.Panels(5).Text = ""
  cmbFBG.AddItem "P = Planillas"
 Else
 If LOC_TIPMOV = 102 Then
 Else
  d_codven.Visible = False
  LBLRUC.Visible = False
  d_Codclie.Visible = False
  lblpersona.Caption = ""
 End If
  lblNumfac.Caption = "Guia"
  lblflete.Caption = ""
  stbEtiqueta.Panels(5).Text = ""
  cmbFBG.AddItem "G = Guia"
 End If
 If LOC_TIPMOV = 10 And LK_FLAG_FACTURACION = "V" Then
     txtvend.Visible = True
     lblvend.Visible = True
 Else
     txtvend.Visible = False
     lblvend.Visible = False
 End If
' temporal = "X"
 cmbFBG.ListIndex = 0
' temporal = ""
 If txtvend.Visible Then
   txtvend.SetFocus
 Else
   If cmbFBG.Visible And cmbFBG.Enabled Then cmbFBG.SetFocus
   If Trim(d_fecha.Caption) = "" Then cmbFBG_KeyPress 13
 End If
 
End If
End Sub


Private Sub TRANS_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 Then
PUB_TIPREG = -10
PUB_CODCIA = LK_CODCIA
Load FrmDatArti
FrmDatArti.Caption = "Mantenimiento de Transportistas"
FrmDatArti.Show 1
TRANSPORTE.Requery
TRANS.Clear
Do Until TRANSPORTE.EOF
    TRANS.AddItem Trim(TRANSPORTE!TRN_NOMBRE) & String(80, " ") & TRANSPORTE!TRN_KEY
    TRANSPORTE.MoveNext
Loop
TRANS.SetFocus
DoEvents

End If
End Sub


Private Sub txtdire_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  FRADIRE.Visible = False
  Exit Sub
End If
If KeyAscii = 13 Then
 Azul txtnum, txtnum
End If
End Sub

Private Sub txtnum_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 27 Then
  FRADIRE.Visible = False
  Exit Sub
End If
If KeyAscii = 13 Then
 TxtZonaTrabajo.SetFocus
 SendKeys "%{UP}"
End If


End Sub

Private Sub txtNumfac_GotFocus()
temporal = txtNumfac.Text
End Sub

Private Sub txtnumfac_KeyPress(KeyAscii As Integer)
'On Error GoTo SALE_X
Dim wven As Integer
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  If val(txtSerie.Text) <= 0 Then
   'Exit Sub
  End If
  If val(txtNumfac.Text) <= 0 Then
   LIMPIA_DOCU
   grid_fac2.Clear
   Exit Sub
  End If
  wflag_docu = ""
  If LOC_TIPMOV = 30 Then
'    txtserie.Locked = True
'    txtserie.Text = 0
    GoSub ORDENES
    Exit Sub
  End If
  If Left(cmbFBG.Text, 1) = "P" And LOC_TIPMOV <> 10 Then
    'txtSerie.Locked = True
    txtSerie.Text = 0
    GoSub PLANILLA
    Exit Sub
  Else
    loc_flag_espera = "A"
    LLENA_CONSULTA
    loc_flag_espera = ""
  End If
  If Trim(wflag_docu) = "" Then
    temporal = txtNumfac.Text
  Else
    txtSerie.Text = tempo_serie
    txtNumfac.Text = temporal
    Azul txtNumfac, txtNumfac
  End If
End If
Exit Sub
PLANILLA:
Dim PS_REP01 As rdoQuery
Dim llave_rep01  As rdoResultset
Dim ws_ingresos As Currency
Dim ws_salidas As Currency
Dim val_ingresos As Currency
Dim val_salidas As Currency
Dim acu_val_ingresos As Currency
Dim acu_val_salidas As Currency
Dim WS_CHEQUE As Currency
d_cheque.Visible = True
d_efectivo.Visible = True
pub_cadena = "SELECT CAA_NUMDOC,CAA_CP, CAA_SERDOC,CAA_SALDO_CAR, CAA_FECHA, CAA_TIPDOC, CAA_CODVEN,CAA_FECHA_VCTO, CAA_CONCEPTO , CAA_CODCLIE,CAA_FECHA_VCTO,CAA_IMPORTE, CAA_SALDO, CAA_NUMSER ,CAA_NUMFAC , CAA_FECHA, CAA_FBG FROM CARACU WHERE CAA_CODCIA = ? AND CAA_NUMPLAN = ? AND CAA_ESTADO <> 'E' ORDER BY CAA_CODCLIE, CAA_FECHA,CAA_NUM_OPER, CAA_SALDO_CAR"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = " "
PS_REP01(1) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

PS_REP01(0) = LK_CODCIA
PS_REP01(1) = val(txtNumfac.Text)
llave_rep01.Requery
d_mensaje.Visible = False
If llave_rep01.EOF = True Then
  grid_fac2.Clear
  LIMPIA_DOCU
  d_mensaje.Visible = True
  GoTo CANCELA
End If
cabe_grid
d_fecha.Caption = Format(llave_rep01!CAA_FECHA, "dd/mm/yyyy")
wven = 0
f1 = 0
WS_BRUTO = 0
grid_fac2.rows = 1
PB.Visible = True
PB.max = llave_rep01.RowCount
PB.Min = 0
PB.Value = 0
Do Until llave_rep01.EOF
   PB.Value = PB.Value + 1
   f1 = f1 + 1
   grid_fac2.rows = grid_fac2.rows + 1
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   pu_cp = "C"
   pu_codclie = llave_rep01!CAA_CODCLIE
   LEER_CLI_LLAVE
   If cli_llave.EOF Then
      MsgBox "Error en Codigo de cliente, NO EXISTE ...", 48, Pub_Titulo
      GoTo CANCELA
   End If
    grid_fac2.TextMatrix(f1, 0) = Trim(cli_llave!CLI_NOMBRE)
    grid_fac2.TextMatrix(f1, 1) = Trim(cli_llave!cli_codclie)
    grid_fac2.TextMatrix(f1, 2) = Trim(llave_rep01!CAA_CODVEN)
   If Trim(llave_rep01!CAA_TIPDOC) = "FA" Or Trim(llave_rep01!CAA_TIPDOC) = "CC" Then
    grid_fac2.TextMatrix(f1, 3) = Trim(llave_rep01!CAA_TIPDOC)
    If llave_rep01!CAA_FBG = "F" Then
      grid_fac2.TextMatrix(f1, 4) = "FAC."
    ElseIf llave_rep01!CAA_FBG = "B" Then
      grid_fac2.TextMatrix(f1, 4) = "BOL."
    ElseIf llave_rep01!CAA_FBG = "G" Then
      grid_fac2.TextMatrix(f1, 4) = "GUIA"
    End If
    grid_fac2.TextMatrix(f1, 5) = llave_rep01!CAa_numser
    grid_fac2.TextMatrix(f1, 6) = llave_rep01!CAa_numfac
    grid_fac2.TextMatrix(f1, 7) = "" 'llave_rep01!CAA_NUMFAC
    grid_fac2.TextMatrix(f1, 8) = Format(llave_rep01!caa_SALDO_car + val(llave_rep01!CAA_IMPORTE * -1), "0.00;(0.00)")
    grid_fac2.TextMatrix(f1, 9) = Format(val(llave_rep01!CAA_IMPORTE * -1), "0.00;(0.00)")
    WS_BRUTO = WS_BRUTO + val(llave_rep01!CAA_IMPORTE * -1)
    grid_fac2.TextMatrix(f1, 10) = llave_rep01!CAA_FECHA_VCTO
  Else
    grid_fac2.TextMatrix(f1, 3) = Trim(llave_rep01!CAA_TIPDOC)
    grid_fac2.TextMatrix(f1, 4) = Trim(llave_rep01!CAA_TIPDOC)
    SQ_OPER = 1
    pu_cp = llave_rep01!CAA_CP
    pu_codclie = cli_llave!cli_codclie
    pu_codcia = LK_CODCIA
    PUB_SERDOC = llave_rep01!caa_serdoc
    PUB_NUMDOC = llave_rep01!CAA_NUMDOC
    PUB_TIPDOC = Trim(llave_rep01!CAA_TIPDOC)
    LEER_CAR_LLAVE
    If Not car_llave.EOF Then grid_fac2.TextMatrix(f1, 11) = car_llave!car_NUM_CHEQUE
    grid_fac2.TextMatrix(f1, 12) = llave_rep01!CAA_IMPORTE * -1
    grid_fac2.TextMatrix(f1, 13) = llave_rep01!caa_concepto
    grid_fac2.TextMatrix(f1, 14) = llave_rep01!CAA_FECHA_VCTO
    WS_CHEQUE = WS_CHEQUE + val(llave_rep01!CAA_IMPORTE * -1)
  End If
  If llave_rep01!CAA_CODVEN <> 0 Then wven = llave_rep01!CAA_CODVEN
  llave_rep01.MoveNext
Loop
frmdocu.d_efectivo.Caption = Format(WS_BRUTO, "0.00;(0.000)")
frmdocu.d_cheque.Caption = Format(WS_CHEQUE, "0.00;(0.000)")
frmdocu.d_saldo.Caption = Format(WS_CHEQUE + WS_BRUTO, "0.00;(0.000)")
SQ_OPER = 1
pu_codcia = LK_CODCIA
PUB_CODVEN = wven
LEER_VEN_LLAVE
d_domicilio.Caption = ""
If Not ven_llave.EOF Then d_domicilio.Caption = Format(wven, "00") + "  " + Trim(ven_llave!VEM_NOMBRE)

lblcheque.Visible = True
lblEfectivo.Visible = True
CmdAnterior.Enabled = True
cmdSiguiente.Enabled = True
PB.Visible = False
TRANS.Visible = False

CANCELA:
Return

ORDENES:
'Dim PS_REP01 As rdoQuery
'Dim llave_rep01  As rdoResultset
'Dim ws_ingresos As Currency
'Dim ws_salidas As Currency
'Dim val_ingresos As Currency
'Dim val_salidas As Currency
'Dim acu_val_ingresos As Currency
'Dim acu_val_salidas As Currency
'Dim WS_CHEQUE As Currency
d_cheque.Visible = True
d_efectivo.Visible = True
pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_TIPMOV=500 AND PED_CODCIA = ? AND PED_NUMSER = ? AND PED_NUMFAC = ? ORDER BY PED_FECHA,PED_NUMSER, PED_NUMFAC,PED_NUMSEC"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = " "
PS_REP01(1) = 0
PS_REP01(2) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

PS_REP01(0) = LK_CODCIA
PS_REP01(1) = txtSerie.Text
PS_REP01(2) = val(txtNumfac.Text)

llave_rep01.Requery
d_mensaje.Visible = False
If llave_rep01.EOF = True Then
  grid_fac2.Clear
  LIMPIA_DOCU
  d_mensaje.Visible = True
  GoTo CANCELA
End If
cabe_grid
d_fecha.Caption = Format(llave_rep01!PED_FECHA, "dd/mm/yyyy")
wven = 0
f1 = 0
WS_BRUTO = 0
grid_fac2.rows = 1
PB.Visible = True
PB.max = llave_rep01.RowCount
PB.Min = 0
PB.Value = 0
pu_codcia = LK_CODCIA
SQ_OPER = 1
pu_cp = "P"
pu_codclie = llave_rep01!PED_CODCLIE
LEER_CLI_LLAVE
If cli_llave.EOF Then
  MsgBox "Rgistro de PROVEEDOR not :" & llave_rep01!PED_CODCLIE, 48, Pub_Titulo
  GoTo CANCELA
End If
PB.Value = PB.Value + 1
d_condicion.Caption = "Orden de Compra"
d_Codclie.Caption = Trim(cli_llave!cli_codclie)
d_nomclie.Caption = Trim(cli_llave!CLI_NOMBRE)
d_domicilio.Caption = Trim(cli_llave!CLI_CASA_DIREC)
d_ruc.Caption = Trim(cli_llave!cli_ruc_esposo)
d_dire.Caption = "      AGENCIA: " & Trim(Nulo_Valors(par_llave!PAR_AGE_EMP))
If Nulo_Valors(cli_llave!CLI_MONEDA) = "D" Then
d_moneda.Caption = "$."
Else
d_moneda.Caption = "S/."
End If
d_usuario.Caption = llave_rep01!PED_CODUSU
d_moneda.Visible = True
If Nulo_Valor0(llave_rep01!PED_MONEDA) = "D" Then
  d_moneda.Caption = "US$"
Else
  d_moneda.Caption = "S/."
End If
cabe_grid
fila = 0
WS_BRUTO = 0
SUB_CANT = 0
subtotal = 0
PUB_DESCTO = 0
grid_fac2.rows = 1
fila = 0
Do Until llave_rep01.EOF
'   PB.Value = PB.Value + 1
   grid_fac2.rows = grid_fac2.rows + 1
   fila = fila + 1
   PUB_KEY = llave_rep01!PED_CODart
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_ART_LLAVE
   If art_LLAVE.EOF Then
      MsgBox "Error Grave en arti..."
   End If
   grid_fac2.TextMatrix(fila, 0) = Trim(art_LLAVE!art_nombre)
   grid_fac2.TextMatrix(fila, 1) = Trim(art_LLAVE!ART_ALTERNO)
   grid_fac2.TextMatrix(fila, 2) = llave_rep01!PED_CANTIDAD / llave_rep01!PED_EQUIV
   grid_fac2.TextMatrix(fila, 3) = llave_rep01!PED_UNIDAD
   grid_fac2.TextMatrix(fila, 4) = Format(llave_rep01!PED_PRECIO, "0.00")
   grid_fac2.TextMatrix(fila, 6) = Format(llave_rep01!ped_DESCTO, "0.0")
   WS_DESCTO = WS_DESCTO + llave_rep01!ped_DESCTO_pre
   subtotal = Format(llave_rep01!PED_PRECIO * (llave_rep01!PED_CANTIDAD / llave_rep01!PED_EQUIV), "0.00")
   subtotal = redondea(subtotal)
   grid_fac2.TextMatrix(fila, 5) = subtotal
   SUB_CANT = SUB_CANT + (llave_rep01!PED_CANTIDAD / llave_rep01!PED_EQUIV)
pasa:
   WS_BRUTO = llave_rep01!PED_BRUTO
   'WS_DESCTO = 0
   WS_IMPTO = llave_rep01!PED_IGV
   WS_GASTOS = 0
   WS_FLETE = 0
   llave_rep01.MoveNext
Loop
   d_subtotal.Caption = WS_BRUTO - WS_DESCTO '- WS_IMPTO +
   d_descto.Caption = WS_DESCTO
   d_flete.Caption = WS_FLETE
   d_impto.Caption = WS_IMPTO
   d_gastos.Caption = WS_GASTOS
   d_neto.Caption = Format(WS_BRUTO + WS_IMPTO - WS_DESCTO + WS_GASTOS, "0.000")
   CmdAnterior.Enabled = True
   cmdSiguiente.Enabled = True
   PB.Visible = False
Exit Sub

Return

SALE_X:
MsgBox Err.Description
End Sub

Public Sub LIMPIA_DOCU()

LBLEXTORNO.Caption = ""
d_Codclie.Caption = ""
d_nomclie.Caption = ""
d_codven.Caption = ""
d_nomven.Caption = ""
d_dias.Caption = ""
d_subtotal.Caption = "0.00"
d_gastos.Caption = "0.00"
d_descto.Caption = "0.000"
d_impto.Caption = "0.000"
d_neto.Caption = "0.000"
d_flete.Caption = "0.000"
d_fechaV.Caption = ""
d_newvcto.Caption = ""
d_saldo.Caption = ""
grid_fac2.Clear
d_dire.Caption = ""
d_ruc.Caption = ""
d_condicion.Caption = ""
d_fecha.Caption = ""
LBLEXTORNO.Visible = False
d_moneda.Caption = ""
d_mensaje.Visible = False
txtdocu.Caption = ""
txtdocu.Visible = False
lblfac.Visible = False
lblcheque.Visible = False
d_mensaje.Visible = False
d_efectivo.Visible = False
d_cheque.Visible = False
d_efectivo.Caption = ""
d_cheque.Caption = ""
lblEfectivo.Visible = False
d_domicilio.Caption = ""
d_usuario.Caption = ""
d_fecha_compra.Caption = ""
FECHA_PART.Text = ""

WGUIA_RELA = ""
LOC_ARROZ = ""
' ICA
If LK_EMP <> "HER" Then
   tguia.Text = ""
End If
End Sub

Public Sub SELE_DOCU()
'Set PSFAR = CN.CreateQuery("", pub_cadena)
'Set far_r = PSFAR.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

End Sub

Private Sub CrearArchivoPlano2(cTipoDocto As String, cSerie As String, cNumero As Double)

Dim oRS As ADODB.Recordset

    LimpiaParametros oCmdEjec
     SIN_CODART = 0
    If grid_fac2.TextMatrix(1, 1) = "" Then SIN_CODART = 1

    If cTipoDocto = "F" Then ' And (cSerie = 1 Or cSerie = 2 Or cSerie = 3 Or cSerie = 4 Or cSerie = 5 Or cSerie = 6) Then
        If LOC_TIPMOV = 10 And SIN_CODART = 1 Then
            oCmdEjec.CommandText = "SP_VENTA_FACTURA_2402_SFS"
        Else
            If Me.d_neto.Caption = 0 Then
            oCmdEjec.CommandText = "SP_VENTA_FACTURA_GRATUITA_SFS"
            Else
            oCmdEjec.CommandText = "SP_VENTA_FACTURA_SFS"
            End If
        End If

    ElseIf cTipoDocto = "B" Then ' And (cSerie = 1 Or cSerie = 2 Or cSerie = 3 Or cSerie = 4 Or cSerie = 5 Or cSerie = 6) Then

        If LOC_TIPMOV = 10 And SIN_CODART = 1 Then
            oCmdEjec.CommandText = "SP_VENTA_BOLETA_2402_SFS"
        Else
            oCmdEjec.CommandText = "SP_VENTA_BOLETA_SFS"
        End If

    ElseIf cTipoDocto = "N" Then
        If LOC_TIPMOV = 97 And SIN_CODART = 1 Then
        oCmdEjec.CommandText = "SP_NOTACREDITO_SFS_CONCEPTO"
        Else
        oCmdEjec.CommandText = "SP_NOTACREDITO_SFS"
        End If
    ElseIf cTipoDocto = "D" And (cSerie = 1 Or cSerie = 2 Or cSerie = 3 Or cSerie = 4 Or cSerie = 5 Or cSerie = 6) Then
        oCmdEjec.CommandText = "SP_NOTADEBITO_SFS"
    ElseIf LK_CODTRA = 1111 Then
        oCmdEjec.CommandText = "SP_COMUNICACION_BAJA_SFS"
        
    End If
    
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@serie", adVarChar, adParamInput, 3, IIf(LK_CODTRA = 1111, PUB_NUMSER_C, cSerie))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numero", adDouble, adParamInput, , IIf(LK_CODTRA = 1111, PUB_NUMFAC_C, cNumero))
    
    
    Set oRS = oCmdEjec.Execute
    
    Dim sCadena As String

    sCadena = ""
    
    Dim obj_FSO     As Object

Dim ArchivoCab  As Object
    Dim ArchivoTri As Object
    Dim ArchivoDet  As Object
    Dim ArchivoLey As Object
    Dim ArchivoAca As Object
    Dim ArchivoPAG As Object
    Dim ArchivoDPA As Object
    Dim ArchivoRTN As Object
    
    Dim sARCHIVOcab As String

    Dim sARCHIVOdet As String
    Dim sARCHIVOtri As String
    Dim sARCHIVOley As String
    Dim sARCHIVOaca As String
    Dim sARCHIVOpag As String
    Dim sARCHIVOdpa As String
    Dim sARCHIVOrtn As String

    Dim sRUC        As String
    
    If LK_CODCIA = "01" Then
     sRUC = Leer_Ini(App.Path & "\config.ini", "RUC", "C:\")
    ElseIf LK_CODCIA = "02" Then
     sRUC = Leer_Ini(App.Path & "\config2.ini", "RUC", "C:\")
    ElseIf LK_CODCIA = "03" Then
     sRUC = Leer_Ini(App.Path & "\config3.ini", "RUC", "C:\")
     Else
     sRUC = Leer_Ini(App.Path & "\config4.ini", "RUC", "C:\")
    End If
     
    sARCHIVOcab = sRUC & "-" & oRS!Nombre + IIf((LOC_TIPMOV = 97 Or LOC_TIPMOV = 98), ".not", IIf(LK_CODTRA = 1111, ".cba", ".cab"))
    sARCHIVOtri = sRUC & "-" & oRS!Nombre + IIf(LK_CODTRA = 2412, ".not", IIf(LK_CODTRA = 1111, ".tri", ".tri"))
    sARCHIVOley = sRUC & "-" & oRS!Nombre + IIf(LK_CODTRA = 2412, ".not", IIf(LK_CODTRA = 1111, ".ley", ".ley"))
    sARCHIVOdet = sRUC & "-" & oRS!Nombre + ".det"
    sARCHIVOaca = sRUC & "-" & oRS!Nombre + ".aca"
    If cTipoDocto = "F" Then
    sARCHIVOpag = sRUC & "-" & oRS!Nombre + ".pag"
    sARCHIVOdpa = sRUC & "-" & oRS!Nombre + ".dpa"
    sARCHIVOrtn = sRUC & "-" & oRS!Nombre + ".rtn"
    End If
    
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")

    'Creamos un archivo con el método CreateTextFile
    If LK_CODCIA = "01" Then
    Set ArchivoCab = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOcab, True)
    Set ArchivoTri = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOtri, True)
    Set ArchivoLey = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOley, True)
    Set ArchivoDet = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOdet, True)
    Set ArchivoAca = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOaca, True)
    If cTipoDocto = "F" Then
      Set ArchivoPAG = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOpag, True)
    End If
    ElseIf LK_CODCIA = "02" Then
    Set ArchivoCab = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOcab, True)
    Set ArchivoTri = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOtri, True)
    Set ArchivoLey = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOley, True)
    Set ArchivoDet = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOdet, True)
    Set ArchivoAca = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOaca, True)
    If cTipoDocto = "F" Then
      Set ArchivoPAG = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config2.ini", "CARPETA", "C:\") + sARCHIVOpag, True)
    End If
    ElseIf LK_CODCIA = "03" Then
    Set ArchivoCab = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOcab, True)
    Set ArchivoTri = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOtri, True)
    Set ArchivoLey = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOley, True)
    Set ArchivoDet = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOdet, True)
    Set ArchivoAca = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOaca, True)
    If cTipoDocto = "F" Then
      Set ArchivoPAG = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config3.ini", "CARPETA", "C:\") + sARCHIVOpag, True)
    End If
    ElseIf LK_CODCIA = "04" Then
    Set ArchivoCab = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOcab, True)
    Set ArchivoTri = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOtri, True)
    Set ArchivoLey = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOley, True)
    Set ArchivoDet = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOdet, True)
    Set ArchivoAca = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOaca, True)
    If cTipoDocto = "F" Then
      Set ArchivoPAG = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config4.ini", "CARPETA", "C:\") + sARCHIVOpag, True)
    End If
    End If
    
    'Set Archivo = obj_FSO.CreateTextFile("C:\" + sARCHIVO, True)
    
    If LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then

        Do While Not oRS.EOF
            'sCadena = sCadena & oRS!fecemision & "|" & oRS!CODMOTIVO & "|" & oRS!DESCMOTIVO & "|" & oRS!TIPODOCAFECTADO & "|" & oRS!NUMDOCAFECTADO & "|" & oRS!TIPDOCUSUARIO & "|" & oRS!NUMDOCUSUARIO & "|" & oRS!CLI1 & "|" & oRS!TIPMONEDA & "|" & oRS!SUMOTROSCARGOS & "|" & oRS!MTOOPERGRAVADAS & "|" & oRS!MTOOPERINAFECTAS & "|" & oRS!MTOOPEREXONERADAS & "|" & oRS!MTOIGV & "|" & oRS!MTOISC & "|" & oRS!MTOOTROSTRIBUTOS & "|" & oRS!MTOIMPVENTA & "|"
            sCadena = sCadena & oRS!CAMPO1 & "|" & oRS!fecemision & "|" & oRS!hORA & "|" & oRS!CAMPO2 & "|" & oRS!CAMPO3 & "|" & oRS!RUC & "|" & oRS!CLI1 & "|" & oRS!TIPMONEDA & "|" & oRS!CODMOTIVO & "|" & Trim(Replace(Replace(oRS!DESCMOTIVO, Chr(10), ""), Chr(13), "")) & "|" & oRS!TIPODOCAFECTADO & "|" & oRS!NUMDOCAFECTADO & "|" & oRS!MTOIGV & "|" & oRS!MTOOPERGRAVADAS & "|" & oRS!MTOIMPVENTA & "|" & oRS!MTOOPERINAFECTAS & "|" & oRS!MTOOTROSTRIBUTOS & "|" & oRS!MTOISC & "|" & oRS!MTOIMPVENTA & "|" & oRS!CAMPO5 & "|" & oRS!CAMPO6 & "|"
            oRS.MoveNext
        Loop
    
    ElseIf LK_CODTRA = 1111 Then
         Do While Not oRS.EOF
            sCadena = sCadena & oRS!FEC_GENERACcION & "|" & oRS!FEC_COMUNICACION & "|" & oRS!TIPDOCBAJA & "|" & oRS!NUMDOCBAJA & "|" & oRS!DESMOTIVOBAJA & "|"
            oRS.MoveNext
        Loop
    Else

        Do While Not oRS.EOF
            'sCadena = sCadena & oRS!TIPOPERACION & "|" & oRS!fecemision & "|" & oRS!hORA & "|" & oRS!FECHAVENC & "|" & oRS!codlocalemisor & "|" & oRS!TIPDOCUSUARIO & "|" & oRS!NUMDOCUSUARIO & "|" & oRS!rznsocialusuario & "|" & oRS!TIPMONEDA & "|" & oRS!MTOIGV & "|" & oRS!MTOOPERGRAVADAS & "|" & oRS!MTOIMPVENTA & "|" & oRS!SUMDSCTOGLOBAL & "|" & oRS!SUMOTROSCARGOS & "|" & oRS!TOTANTICIPOS & "|" & oRS!IMPTOTALVENTA & "|" & oRS!UBL & "|" & oRS!CUSTOMDOC & "|"
            sCadena = sCadena & oRS!TIPOPERACION & "|" & oRS!fecemision & "|" & oRS!hORA & "|" & oRS!FECHAVENC & "|" & oRS!codlocalemisor & "|" & oRS!TIPDOCUSUARIO & "|" & oRS!NUMDOCUSUARIO & "|" & oRS!rznsocialusuario & "|" & oRS!TIPMONEDA & "|" & oRS!MTOIGV & "|" & oRS!MTOOPERGRAVADAS & "|" & oRS!MTOIMPVENTA & "|" & oRS!SUMDSCTOGLOBAL & "|" & oRS!SUMOTROSCARGOS & "|" & oRS!TOTANTICIPOS & "|" & oRS!IMPTOTALVENTA & "|" & oRS!UBL & "|" & oRS!CUSTOMDOC & "|"
            oRS.MoveNext
        Loop

    End If
   
    'Escribimos lineas
    ArchivoCab.WriteLine sCadena
    
    'Cerramos el fichero
    ArchivoCab.Close
    Set ArchivoCab = Nothing
    
     'DIRECCION
    oRS.MoveFirst
    sCadena = ""
    Do While Not oRS.EOF
        'sCadena = sCadena & oRS!ACA1 & "|" & oRS!ACA2 & "|" & oRS!ACA3 & "|" & oRS!ACA4 & "|" & oRS!PAIS & "|" & oRS!UBIGEO & "|" & oRS!dir & "|" & oRS!PAIS1 & "|" & oRS!UBIGEO1 & "|" & oRS!dir1 & "|"
        sCadena = sCadena & oRS!ACA1 & "|" & oRS!ACA2 & "|" & oRS!ACA3 & "|" & oRS!ACA4 & "|" & oRS!ACA5 & "|" & oRS!PAIS & "|" & oRS!UBIGEO & "|" & oRS!dir & "|" & oRS!PAIS1 & "|" & oRS!UBIGEO1 & "|" & oRS!dir1 & "|"
        oRS.MoveNext
    Loop
    
    'Escribimos LINEAS
    ArchivoAca.WriteLine sCadena
    
    'Cerramos el fichero
    ArchivoAca.Close
    Set ArchivoAca = Nothing
   
    Dim oRSdet As ADODB.Recordset

    Set oRSdet = oRS.NextRecordset
   
    sCadena = ""
    Dim c As Integer
    c = 1

    If LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then

        Do While Not oRSdet.EOF
         
            'sCadena = sCadena & oRSdet!CODUNIDADMEDIDA & "|" & oRSdet!CTDUNIDADITEM & "|" & oRSdet!CODPRODUCTO & "|" & oRSdet!CODPRODUCTOSUNAT & "|" & RTrim(oRSdet!desitem) & "|" & oRSdet!MTOVALORUNITARIO & "|" & oRSdet!MTOIGVITEM & "|" & oRSdet!CODTIPTRIBUTOIGV & "|" & oRSdet!MTOIGVITEM1 & "|" & oRSdet!NOMTRIBITEM & "|" & oRSdet!CODTIPTRIBUTOITEM & "|" & oRSdet!TIPAFEIGV & "|" & FormatNumber(oRSdet!PORCIGV, 2) & "|" & oRSdet!CODISC & "|" & oRSdet!CODOTROITEM & "|" & oRSdet!GRATUITO & "|"
            'sCadena = sCadena & oRSdet!CODUNIDADMEDIDA & "|" & oRSdet!CTDUNIDADITEM & "|" & oRSdet!CODPRODUCTO & "|" & oRSdet!CODPRODUCTOSUNAT & "|" & Trim(Replace(Replace(oRSdet!DESITEM, Chr(10), ""), Chr(13), "")) & "|" & oRSdet!MTOVALORUNITARIO & "|" & oRSdet!MTOIGVITEM & "|" & oRSdet!CODTIPTRIBUTOIGV & "|" & oRSdet!MTOIGVITEM & "|" & oRSdet!MTOVALOR2 & "|IGV|VAT|" & oRSdet!TIPAFEIGV & "|" & oRSdet!PORCIGV & "|" & oRSdet!CEROS & "|" & oRSdet!MTOISCITEM & "|" & oRSdet!MTOISCITEM & "|||" & oRSdet!TIPSISISC & "|" & oRSdet!GRATUITO & "|-|" & oRSdet!MTOISCITEM & "|" & oRSdet!MTOISCITEM & "|||0|" & (oRSdet!MTOIGVITEM + oRSdet!MTOVALOR2) & "|" & oRSdet!MTOVALOR2 & "|0.00|"
            sCadena = sCadena & oRSdet!CODUNIDADMEDIDA & "|" & oRSdet!CTDUNIDADITEM & "|" & oRSdet!CODPRODUCTO & "|" & oRSdet!CODPRODUCTOSUNAT & "|" & Trim(oRSdet!DESITEM) & _
           "|" & oRSdet!MTOVALORUNITARIO & "|" & oRSdet!MTOIGVITEM & "|" & oRSdet!CODTIPTRIBUTOIGV & "|" & oRSdet!MTOIGVITEM1 & "|" & oRSdet!BASEIMPIGV & "|" & _
           oRSdet!NOMTRIBITEM & "|" & oRSdet!CODTIPTRIBUTOITEM & "|" & oRSdet!TIPAFEIGV & "|" & FormatNumber(oRSdet!PORCIGV, 2) & "|" & oRSdet!CODISC & "|" & oRSdet!MONTOISC & _
           "|" & oRSdet!BASEIMPONIBLEISC & "|" & oRSdet!NOMBRETRIBITEM & "|" & oRSdet!CODTRIBITEM & "|" & oRSdet!CODSISISC & "|" & oRSdet!PORCISC & "|" & oRSdet!CODTRIBOTO & _
           "|" & oRSdet!MONTOTRIBOTO & "|" & oRSdet!BASEIMPONIBLEOTO & "|" & oRSdet!NOMBRETRIBOTO & "|" & oRSdet!TIPSISISC & "|" & oRSdet!PORCOTO & "|" & oRSdet!CODIGOICBPER & _
           "|" & oRSdet!IMPORTEICBPER & "|" & oRSdet!CANTIDADICBPER & "|" & oRSdet!TITULOICBPER & "|" & oRSdet!IDEICBPER & "|" & oRSdet!MONTOICBPER & "|" & _
           oRSdet!PRECIOVTAUNITARIO & "|" & oRSdet!VALORVTAXITEM & "|" & oRSdet!GRATUITO & "|"
            If c < oRSdet.RecordCount Then
                sCadena = sCadena + vbCrLf
            End If
             c = c + 1
            oRSdet.MoveNext
            
        Loop

    ElseIf LK_CODTRA <> 1111 Then
    

        Do While Not oRSdet.EOF
       
           ' sCadena = sCadena & oRSdet!CODUNIDADMEDIDA & "|" & oRSdet!CTDUNIDADITEM & "|" & oRSdet!CODPRODUCTO & "|" & oRSdet!CODPRODUCTOSUNAT & "|" & oRSdet!DESITEM & "|" & oRSdet!MTOVALORUNITARIO & "|" & oRSdet!MTODSCTOITEM & "|" & oRSdet!MTOIGVITEM & "|" & oRSdet!TIPAFEIGV & "|" & oRSdet!MTOISCITEM & "|" & oRSdet!TIPSISISC & "|" & oRSdet!MTOPRECIOVENTAITEM & "|" & oRSdet!MTOVALORVENTAITEM & "|"
          ' sCadena = sCadena & oRSdet!CODUNIDADMEDIDA & "|" & oRSdet!CTDUNIDADITEM & "|" & oRSdet!CODPRODUCTO & "|" & oRSdet!CODPRODUCTOSUNAT & "|" & Trim(Replace(Replace(oRSdet!DESITEM, Chr(10), ""), Chr(13), "")) & "|" & oRSdet!MTOVALORUNITARIO & "|" & oRSdet!MTOIGVITEM & "|" & oRSdet!CODTIPTRIBUTOIGV & "|" & oRSdet!MTOIGVITEM1 & "|" & oRSdet!BASEIMPIGV & "|" & oRSdet!NOMTRIBITEM & "|" & oRSdet!CODTIPTRIBUTOITEM & "|" & oRSdet!TIPAFEIGV & "|" & FormatNumber(oRSdet!PORCIGV, 2) & "|" & oRSdet!CODISC & "|" & oRSdet!MONTOISC & "|" & oRSdet!BASEIMPONIBLEISC & "|" & oRSdet!NOMBRETRIBITEM & "|" & oRSdet!CODTRIBITEM & "|" & oRSdet!CODSISISC & "|" & oRSdet!PORCISC & "|" & oRSdet!CODTRIBOTO & "|" & oRSdet!MONTOTRIBOTO & "|" & oRSdet!BASEIMPONIBLEOTO & "|" & oRSdet!NOMBRETRIBOTO & "|" & oRSdet!TIPSISISC & "|" & oRSdet!PORCOTO & "|" & oRSdet!PRECIOVTAUNITARIO & "|" & oRSdet!VALORVTAXITEM & "|" & oRSdet!GRATUITO & "|"
           sCadena = sCadena & oRSdet!CODUNIDADMEDIDA & "|" & oRSdet!CTDUNIDADITEM & "|" & oRSdet!CODPRODUCTO & "|" & oRSdet!CODPRODUCTOSUNAT & "|" & Trim(oRSdet!DESITEM) & _
           "|" & oRSdet!MTOVALORUNITARIO & "|" & oRSdet!MTOIGVITEM & "|" & oRSdet!CODTIPTRIBUTOIGV & "|" & oRSdet!MTOIGVITEM1 & "|" & oRSdet!BASEIMPIGV & "|" & _
           oRSdet!NOMTRIBITEM & "|" & oRSdet!CODTIPTRIBUTOITEM & "|" & oRSdet!TIPAFEIGV & "|" & FormatNumber(oRSdet!PORCIGV, 2) & "|" & oRSdet!CODISC & "|" & oRSdet!MONTOISC & _
           "|" & oRSdet!BASEIMPONIBLEISC & "|" & oRSdet!NOMBRETRIBITEM & "|" & oRSdet!CODTRIBITEM & "|" & oRSdet!CODSISISC & "|" & oRSdet!PORCISC & "|" & oRSdet!CODTRIBOTO & _
           "|" & oRSdet!MONTOTRIBOTO & "|" & oRSdet!BASEIMPONIBLEOTO & "|" & oRSdet!NOMBRETRIBOTO & "|" & oRSdet!TIPSISISC & "|" & oRSdet!PORCOTO & "|" & oRSdet!CODIGOICBPER & _
           "|" & oRSdet!IMPORTEICBPER & "|" & oRSdet!CANTIDADICBPER & "|" & oRSdet!TITULOICBPER & "|" & oRSdet!IDEICBPER & "|" & oRSdet!MONTOICBPER & "|" & _
           oRSdet!PRECIOVTAUNITARIO & "|" & oRSdet!VALORVTAXITEM & "|" & oRSdet!GRATUITO & "|"
            If c < oRSdet.RecordCount Then
                sCadena = sCadena + vbCrLf
            End If
             c = c + 1
            oRSdet.MoveNext
             
        Loop

    End If

    'Escribimos lineas
    If LK_CODTRA <> 1111 Then
        ArchivoDet.WriteLine sCadena
        
         'Cerramos el fichero
        ArchivoDet.Close
        Set ArchivoDet = Nothing
        
        Dim orsTri As ADODB.Recordset
        Set orsTri = oRS.NextRecordset
        
        sCadena = ""
        c = 1
        'ARCIVO .TRI
        Do While Not orsTri.EOF
        'sCadena = sCadena & orsTri!Codigo & "|" & orsTri!Nombre & "|" & orsTri!cod & "|" & orsTri!BASEIMPONIBLE & "|" & orsTri!TRIBUTO & "|"
        sCadena = sCadena & orsTri!Codigo & "|" & orsTri!Nombre & "|" & orsTri!cod & "|" & orsTri!BASEIMPONIBLE & "|" & orsTri!TRIBUTO & "|"
        If c < orsTri.RecordCount Then
            sCadena = sCadena & vbCrLf
        End If
        c = c + 1
            orsTri.MoveNext
        Loop
        
        
         ArchivoTri.WriteLine sCadena
        
         'Cerramos el fichero
        ArchivoTri.Close
        Set ArchivoTri = Nothing
        
        Dim orsLey As ADODB.Recordset
        Set orsLey = oRS.NextRecordset
        Dim moneda As String
        
        If d_moneda.Caption = "S/." Then
        moneda = "S"
        Else
        moneda = "D"
        End If
        
        
        c = 1
        sCadena = ""
        Do While Not orsLey.EOF
            sCadena = sCadena & orsLey!cod & "|" & Trim(CONVER_LETRAS(Me.d_neto.Caption, moneda)) & "|"
            If c < orsLey.RecordCount Then
                sCadena = sCadena & vbCrLf
            End If
            c = c + 1
            orsLey.MoveNext
        Loop
        
        ArchivoLey.WriteLine sCadena
        ArchivoLey.Close
        Set ArchivoLey = Nothing
        
      Dim xFormaPago As String
    If cTipoDocto = "F" Then
            'PAG
            Dim orsPAG As ADODB.Recordset
            Set orsPAG = oRS.NextRecordset
            
            c = 1
            sCadena = ""
            Do While Not orsPAG.EOF
                xFormaPago = orsPAG!formaPAGO
                sCadena = sCadena & orsPAG!formaPAGO & "|" & orsPAG!pendientepago & "|" & orsPAG!TIPMONEDA & "|"
                If c < orsPAG.RecordCount Then
                    sCadena = sCadena & vbCrLf
                End If
                c = c + 1
                orsPAG.MoveNext
            Loop
            
            ArchivoPAG.WriteLine sCadena
            ArchivoPAG.Close
            Set ArchivoPAG = Nothing
            
            'DPA
            Dim orsDPA As ADODB.Recordset
            Set orsDPA = oRS.NextRecordset
            If UCase(xFormaPago) = "CREDITO" Or UCase(xFormaPago) = "CRÉDITO" Then
                Set ArchivoDPA = obj_FSO.CreateTextFile(Leer_Ini(App.Path & "\config.ini", "CARPETA", "C:\") + sARCHIVOdpa, True)
               
                
                c = 1
                sCadena = ""
                Do While Not orsDPA.EOF
                    sCadena = sCadena & orsDPA!cuotapago & "|" & orsDPA!fechavcto & "|" & orsDPA!TIPMONEDA & "|"
                    If c < orsDPA.RecordCount Then
                        sCadena = sCadena & vbCrLf
                    End If
                    c = c + 1
                    orsDPA.MoveNext
                Loop
                
                ArchivoDPA.WriteLine sCadena
                ArchivoDPA.Close
                Set ArchivoDPA = Nothing
            End If
             'RTN
'            Dim orsRTN As ADODB.Recordset
'            Set orsRTN = oRS.NextRecordset
'
'            c = 1
'            sCadena = ""
'            Do While Not orsRTN.EOF
'                sCadena = sCadena & orsRTN!impoperacion & "|" & orsRTN!porretencion & "|" & orsRTN!impretencion & "|"
'                If c < orsRTN.RecordCount Then
'                    sCadena = sCadena & vbCrLf
'                End If
'                c = c + 1
'                orsRTN.MoveNext
'            Loop
'
'            ArchivoRTN.WriteLine sCadena
'            ArchivoRTN.Close
'            Set ArchivoRTN = Nothing
        End If
    
    End If
    
   
    
    Set obj_FSO = Nothing
End Sub




Public Sub LEER_FAR_CONSUL()
PSFAR_CONSUL.rdoParameters(0) = PU_TIPMOV
PSFAR_CONSUL.rdoParameters(1) = pu_codcia
PSFAR_CONSUL.rdoParameters(2) = PU_FBG
PSFAR_CONSUL.rdoParameters(3) = PU_NUMSER
PSFAR_CONSUL.rdoParameters(4) = pu_cp

'If LOC_TIPMOV = 97 And (Right(cmbFBG.Text, 1) = "P" Or Right(cmbFBG.Text, 1) = "C") Then
' PSFAR_CONSUL.rdoParameters(4) = Right(cmbFBG.Text, 1)
'End If
far_consul.Requery


End Sub
Public Sub LEER_CAR_CONSUL()
PSCAR_CONSUL.rdoParameters(0) = pu_codcia
PSCAR_CONSUL.rdoParameters(1) = pu_cp
PSCAR_CONSUL.rdoParameters(2) = pu_codclie
PSCAR_CONSUL.rdoParameters(3) = PU_FBG
PSCAR_CONSUL.rdoParameters(4) = PU_NUMSER
PSCAR_CONSUL.rdoParameters(5) = PU_NUMFAC
car_consul.Requery
End Sub

Private Sub txtserie_GotFocus()
txtSerie.Text = Trim(txtSerie.Text)
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 If txtNumfac.Enabled Then
  txtNumfac.SetFocus
 End If
End If
End Sub
Public Sub LLENADOS(cont As ComboBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
        If tab_mayor!TAB_CODART = 1 Then cont.AddItem tab_mayor!tab_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
        tab_mayor.MoveNext
    Loop
End Sub

Public Function REP_CONSUL() As Integer
Dim WMONEDA As String * 1
Dim Xx As String * 1
Dim wser As String * 3
Dim WSRUTA As String
Dim indice As Integer
Dim wm As Integer
Dim llave_rep01 As rdoResultset
Dim PS_REP01 As rdoQuery
Dim i As Integer
Dim valor
Dim loc_xl As Object
Dim loc_codtra As Integer
Dim wRuta As String
Dim WSNUMDOC As String
Dim numero_device As Integer
If LOC_TIPMOV = 3 Or LOC_TIPMOV = 100 Or LOC_TIPMOV = 101 Or LOC_TIPMOV = 10 Then
    SQ_OPER = 2
    PUB_CODCIA = LK_CODCIA
    If LK_FLAG_FACTURACION = "A" Then
        PUB_CODVEN = 1
    ElseIf LK_FLAG_FACTURACION = "V" Then
        PUB_CODVEN = 1
    End If
    LEER_PAR_LLAVE
    If pac_llave.EOF Then
       MsgBox "No se ha definido archivos de Impresión", 48, Pub_Titulo
       Exit Function
    End If
End If

'If LK_EMP = "HER" Then
'  wRuta = "C:\ADMIN\STANDAR\"
'Else
If LK_EMP_PTO = "A" Then
  wRuta = PUB_RUTA_OTRO & "PTOVTA\"
Else
  wRuta = PUB_RUTA_OTRO
End If
If Trim(d_moneda.Caption) = "S/." Then
 WMONEDA = "S"
Else
 WMONEDA = "D"
End If

'End If
  

    frmdocu.Reportes.Connect = PUB_ODBC
    If frmdocu.imp.Value = 1 Then
      frmdocu.Reportes.Destination = crptToPrinter
    Else
      frmdocu.Reportes.Destination = crptToWindow  '= crptToPrinter
    End If
    frmdocu.Reportes.WindowLeft = 2
    frmdocu.Reportes.WindowTop = 70
    frmdocu.Reportes.WindowWidth = 635
    frmdocu.Reportes.WindowHeight = 390
    frmdocu.Reportes.Formulas(1) = ""
    frmdocu.d_neto.Refresh
    PUB_NETO = val(frmdocu.d_neto.Caption)
    PUB_FECHA = frmdocu.d_fecha.Caption
    PU_NUMSER = val((frmdocu.txtSerie.Text))
    PU_NUMFAC = val((frmdocu.txtNumfac.Text))
    If LK_EMP = "PIU" Then
       frmdocu.Reportes.Formulas(1) = "SON=  ' " & CONVER_LETRAS(PUB_NETO, WMONEDA) & "'"
    Else
       frmdocu.Reportes.Formulas(1) = "SON=  ' " & CONVER_LETRAS(PUB_NETO, WMONEDA) & "'"  'quitado SON GTS
    End If
    If PUB_NETO <> val(frmdocu.d_neto.Caption) Then
      MsgBox "Espere....!!!", 48, Pub_Titulo
      Exit Function
    End If
    LOC_NUMFAC_FIN = PU_NUMFAC
    Reportes.Formulas(8) = ""
    Reportes.Formulas(9) = ""
    Reportes.Formulas(10) = ""
    Reportes.Formulas(11) = ""
    Reportes.Formulas(12) = ""
    Reportes.Formulas(13) = ""
    Reportes.Formulas(14) = ""
    Reportes.Formulas(15) = ""
    Reportes.Formulas(16) = ""
    Reportes.Formulas(17) = ""
    Reportes.Formulas(18) = ""
    Reportes.Formulas(19) = ""
    Reportes.Formulas(20) = ""
    Reportes.Formulas(21) = ""
    Reportes.Formulas(22) = ""
    Reportes.Formulas(23) = ""
    Reportes.Formulas(24) = ""
    Reportes.Formulas(25) = ""
    Reportes.Formulas(26) = ""
    Reportes.Formulas(27) = ""
    Reportes.Formulas(28) = ""

    
    
    If LOC_TIPMOV = 3 Or LOC_TIPMOV = 5 Or LOC_TIPMOV = 6 Or LOC_TIPMOV = 100 Or LOC_TIPMOV = 101 Then
        frmdocu.Reportes.WindowTitle = "GUIA DE REMISION  :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
        If LOC_TIPMOV = 3 Or LOC_TIPMOV = 100 Or LOC_TIPMOV = 101 Then
          frmdocu.Reportes.ReportFileName = wRuta + Trim(pac_llave!PAC_ARCHI_GUIA)  ' "FACGUIA.RPT"
        Else
          frmdocu.Reportes.ReportFileName = wRuta + "GUIAR3.RPT"
        End If
        pub_cadena = "{FACART.FAR_ESTADO} <> 'E' AND {FACART.FAR_TIPMOV} = " & LOC_TIPMOV & " AND {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND  ( {FACART.FAR_NUMFAC} >= " & PU_NUMFAC & " and {FACART.FAR_NUMFAC} <= " & LOC_NUMFAC_FIN & ") AND {FACART.FAR_NUMSER} = '" & Trim(txtSerie.Text) & "'  "
        'Debug.Print pub_cadena
        frmdocu.Reportes.Formulas(1) = ""
        If chetrans.Value = 1 Then
             PS_TRA(0) = val(Right(TRANS.Text, 3))
             llave_trans.Requery
             Reportes.Formulas(12) = "TRN_NOMBRE    =  '" & llave_trans!TRN_NOMBRE & "'"
             Reportes.Formulas(13) = "TRN_DIRECCION =  '" & llave_trans!TRN_DIRECCION & "'"
             Reportes.Formulas(14) = "TRN_RUC       =  '" & llave_trans!TRN_RUC & "'"
             Reportes.Formulas(15) = "TRN_DNI       =  '" & llave_trans!TRN_DNI & "'"
             Reportes.Formulas(16) = "TRN_PLACA     =  '" & llave_trans!TRN_PLACA & "'"
             Reportes.Formulas(16) = "TRN_CHOFER    =  '" & llave_trans!TRN_CHOFER & "'"
        End If
        If LOC_TIPMOV = 100 Or LOC_TIPMOV = 3 Then GoTo PASA_OP
        GoTo pasa_todo
    End If
    If LOC_TIPMOV = 20 Then
        frmdocu.Reportes.WindowTitle = "KARDEX Nº :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
        pub_mensaje = "Inventario Valorado (Si), Inventario en Unidades (No) "
        Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
        If Pub_Respuesta = vbYes Then
            frmdocu.Reportes.ReportFileName = wRuta + "NOTAING.RPT"
        Else
            frmdocu.Reportes.ReportFileName = wRuta + "NOTAINV.RPT"
        End If
        wser = PU_NUMSER
        pub_cadena = "{FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_NUMSER}= '" & wser & "'  AND {FACART.FAR_NUMFAC} = " & PU_NUMFAC
        frmdocu.Reportes.Formulas(1) = ""
        Reportes.Formulas(1) = "CIA=  '" & Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))) & "'"
        GoTo pasa_todo
    End If
    If LOC_TIPMOV = 75 Or LOC_TIPMOV = 93 Or LOC_TIPMOV = 102 Then
        wser = PU_NUMSER
        frmdocu.Reportes.WindowTitle = "KARDEX Nº :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
        If LOC_TIPMOV = 93 Then
        frmdocu.Reportes.ReportFileName = wRuta + "CAMBP.RPT"
        End If
        If LOC_TIPMOV = 75 Then
          frmdocu.Reportes.ReportFileName = wRuta + "DEF01.RPT"
        End If
        If LOC_TIPMOV = 102 Then
          frmdocu.Reportes.ReportFileName = wRuta + "CAMBPRO.RPT"
          Reportes.Formulas(12) = "DOCUMENTO=  '" & DOC102 & "'"
        End If
        pub_cadena = "{FACART.FAR_TIPMOV} = " & LOC_TIPMOV & " AND {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_NUMSER}= '" & wser & "'  AND {FACART.FAR_NUMFAC} = " & PU_NUMFAC
        frmdocu.Reportes.Formulas(1) = ""
        GoTo pasa_todo
    End If
    If LOC_TIPMOV = 99 Then
        wser = PU_NUMSER
        frmdocu.Reportes.WindowTitle = "KARDEX Nº :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
        frmdocu.Reportes.ReportFileName = wRuta + "VOCCM.RPT"
        pub_cadena = "{ALLOG.ALL_CODCIA} = '" & LK_CODCIA & "' AND {ALLOG.ALL_NUMSER}= '" & wser & "'  AND {ALLOG.ALL_NUMFAC} = " & PU_NUMFAC
        frmdocu.Reportes.Formulas(1) = ""
        GoTo pasa_todo
    End If
    If LOC_TIPMOV = 10 Or LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
        SQ_OPER = 1 'Dato de Operacion al leer_far_llave --JC
        LEER_FAR_LLAVE 'Leo los datos de facart --JC
      If Left(frmdocu.cmbFBG.Text, 1) = "B" Then
        frmdocu.Reportes.WindowTitle = "BOLETA  :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000") & " al " & Format(LOC_NUMFAC_FIN, "0000000")
        If LK_EMP = "PIU" Or LK_EMP = "PAR" Then
          frmdocu.Reportes.WindowTitle = "BOLETA  :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
        End If
        'CONDICION PARA IMPRIMIR: SI ES PARA CONSUMO VA A OTRO REPORTE
        'SINO ES EL NOMBRE DEL DATO DE LA COLUMNA PAC_ARCHI_B DE LA TABLA PARARC--JC
        
        
        If far_llave!far_JABAS = 0 Then 'por detalle
            frmdocu.Reportes.ReportFileName = wRuta + Trim(pac_llave!PAC_ARCHI_B)
        Else
            frmdocu.Reportes.ReportFileName = wRuta + "BOLCONSUMO.rpt"
        End If
        '**********************************
        'frmdocu.Reportes.ReportFileName = wRuta + Trim(pac_llave!PAC_ARCHI_B)  '"CLIBOL.RPT"
        
       If SIN_CODART = 1 Then frmdocu.Reportes.ReportFileName = wRuta + "BOL002.RPT"  ' agregado gts para imprimir por concepto
        If LOC_ARROZ = "A" Then frmdocu.Reportes.ReportFileName = wRuta + "CLIBOLIE.RPT"
      ElseIf Left(frmdocu.cmbFBG.Text, 1) = "F" Then
        frmdocu.Reportes.WindowTitle = "FACTURA : " & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000") & " al " & Format(LOC_NUMFAC_FIN, "0000000")
        If LK_EMP = "PIU" Or LK_EMP = "PAR" Then
          frmdocu.Reportes.WindowTitle = "FACTURA : " & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
        End If
        'CONDICION PARA IMPRIMIR: SI ES PARA CONSUMO VA A OTRO REPORTE
        'SINO ES EL NOMBRE DEL DATO DE LA COLUMNA PAC_ARCHI_F DE LA TABLA PARARC--JC
        
        
        
        If far_llave!far_JABAS = 0 Then 'por detalle
            frmdocu.Reportes.ReportFileName = wRuta + Trim(pac_llave!PAC_ARCHI_F)
        Else
            frmdocu.Reportes.ReportFileName = wRuta + "FACCONSUMO.rpt"
        End If
        '**********************************
        'frmDocu.Reportes.ReportFileName = wRuta + Trim(pac_llave!PAC_ARCHI_F)
        If LOC_ARROZ = "A" Then frmdocu.Reportes.ReportFileName = wRuta + "CLIFACIE.RPT"
        If SIN_CODART = 1 Then frmdocu.Reportes.ReportFileName = wRuta + "FAC002.RPT"  ' agregado gts para imprimir por concepto
        
       ElseIf Left(frmdocu.cmbFBG.Text, 1) = "T" Then
        frmdocu.Reportes.WindowTitle = "TICKET : " & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000") & " al " & Format(LOC_NUMFAC_FIN, "0000000")
        If LK_EMP = "PIU" Or LK_EMP = "PAR" Then
          frmdocu.Reportes.WindowTitle = "TICKET : " & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
        End If
        'CONDICION PARA IMPRIMIR: SI ES PARA CONSUMO VA A OTRO REPORTE
        'SINO ES EL NOMBRE DEL DATO DE LA COLUMNA PAC_ARCHI_F DE LA TABLA PARARC--JC
        
        
        
        If far_llave!far_JABAS = 0 Then 'por detalle
            frmdocu.Reportes.ReportFileName = wRuta + "TICKETDETALLE.rpt"
        Else
            frmdocu.Reportes.ReportFileName = wRuta + "TICKETCONSUMO.rpt"
        End If
        '**********************************
        'frmDocu.Reportes.ReportFileName = wRuta + Trim(pac_llave!PAC_ARCHI_F)
        If LOC_ARROZ = "A" Then frmdocu.Reportes.ReportFileName = wRuta + "CLIFACIE.RPT"
        If SIN_CODART = 1 Then frmdocu.Reportes.ReportFileName = wRuta + "FAC002.RPT"  ' agregado gts para imprimir por concepto
      ElseIf Left(frmdocu.cmbFBG.Text, 1) = "G" Then
        frmdocu.Reportes.WindowTitle = " GUIA   : " & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000") & " al " & Format(LOC_NUMFAC_FIN, "0000000")
        frmdocu.Reportes.ReportFileName = wRuta + Trim(pac_llave!PAC_ARCHI_G)
      ElseIf Left(frmdocu.cmbFBG.Text, 1) = "N" Then
        frmdocu.Reportes.WindowTitle = " N. CREDITO  : " & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000") & " al " & Format(LOC_NUMFAC_FIN, "0000000")
        If Trim(grid_fac2.TextMatrix(1, 1)) = "" Then
         frmdocu.Reportes.ReportFileName = wRuta + "NCREDV.RPT"
         Else
         frmdocu.Reportes.ReportFileName = wRuta + "NCRED.RPT"
        End If
      ElseIf Left(frmdocu.cmbFBG.Text, 1) = "D" Then
        frmdocu.Reportes.WindowTitle = " N. DEBITO  : " & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
        If Trim(grid_fac2.TextMatrix(1, 1)) = "" Then
         frmdocu.Reportes.ReportFileName = wRuta + Trim(pac_llave!PAC_ARCHI_NDV)
        Else
         frmdocu.Reportes.ReportFileName = wRuta + Trim(pac_llave!PAC_ARCHI_ND)
        End If
      ElseIf Left(Trim(frmdocu.cmbFBG.Text), 1) = "P" Then
         frmdocu.Reportes.WindowTitle = "PEDIDO : " & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "00000000")
         'CONDICION PARA IMPRIMIR: SI ES DETALLE O CONSUMO
         If far_llave!far_JABAS = 0 Then 'por detalle
            frmdocu.Reportes.ReportFileName = wRuta + Trim(pac_llave!PAC_ARCHI_G)
        Else
            frmdocu.Reportes.ReportFileName = wRuta + "PEDCONSUMO.rpt"
        End If
        '**********************************
      End If
    End If
    wser = PU_NUMSER
    If Left(frmdocu.cmbFBG.Text, 1) = "N" Then
       pub_cadena = "{FACART.FAR_ESTADO} <> 'E' AND {FACART.FAR_TIPMOV} = 97 AND {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_FBG} = '" & Left(frmdocu.cmbFBG.Text, 1) & "' AND {FACART.FAR_NUMSER}= '" & wser & "' AND ( {FACART.FAR_NUMFAC} >= " & PU_NUMFAC & " and {FACART.FAR_NUMFAC} <= " & LOC_NUMFAC_FIN & ")"
    ElseIf Left(frmdocu.cmbFBG.Text, 1) = "D" Then
       pub_cadena = "{FACART.FAR_ESTADO} <> 'E' AND {FACART.FAR_TIPMOV} = 98 AND {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_FBG} = '" & Left(frmdocu.cmbFBG.Text, 1) & "' AND {FACART.FAR_NUMSER}= '" & wser & "' AND ( {FACART.FAR_NUMFAC} >= " & PU_NUMFAC & " and {FACART.FAR_NUMFAC} <= " & LOC_NUMFAC_FIN & ")"
    Else
       pub_cadena = "{FACART.FAR_ESTADO} <> 'E' AND {FACART.FAR_TIPMOV} = 10 AND {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_FBG} = '" & Left(frmdocu.cmbFBG.Text, 1) & "' AND {FACART.FAR_NUMSER}= '" & wser & "' AND ( {FACART.FAR_NUMFAC} >= " & PU_NUMFAC & " and {FACART.FAR_NUMFAC} <= " & LOC_NUMFAC_FIN & ")"
     End If
    Reportes.Formulas(12) = ""
    Reportes.Formulas(13) = ""
    Reportes.Formulas(14) = ""
    Reportes.Formulas(15) = ""
    Reportes.Formulas(16) = ""
     ' ICA
    If cherela.Value = 1 And LK_EMP = "HER" Then
       Reportes.Formulas(10) = "GUIA=  '" & Trim(tguia.Text) & "'"
    End If
    If Left(frmdocu.cmbFBG.Text, 1) = "G" Then GoTo PASA_OP
    If LK_EMP = "PIU" Or LK_EMP = "PAR" Or LK_EMP = "3AA" Then
       Reportes.Formulas(17) = ""
       Reportes.Formulas(18) = ""
       Reportes.Formulas(19) = ""
       Reportes.Formulas(20) = ""
       pub_mensaje = "Desea Imprimir la " & Trim(frmdocu.Reportes.WindowTitle) & "   ¿Desea Continuar... ?"
       Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
       If Pub_Respuesta = vbYes Then
         frmdocu.Reportes.SelectionFormula = pub_cadena
         Reportes.Formulas(10) = ""
         Reportes.Formulas(11) = ""
         Reportes.Formulas(12) = ""
         Reportes.Formulas(13) = ""
         Reportes.Formulas(14) = ""
         Reportes.Formulas(15) = ""
         Reportes.Formulas(16) = ""
           Reportes.Formulas(21) = ""
           Reportes.Formulas(22) = ""
           Reportes.Formulas(23) = ""
           Reportes.Formulas(24) = ""
           Reportes.Formulas(25) = ""
           Reportes.Formulas(26) = ""
           Reportes.Formulas(27) = ""
         If cherela.Value = 1 Then
            Reportes.Formulas(10) = "GUIA=  '" & Trim(tguia.Text) & "'"
         End If
         On Error GoTo accion
         'Debug.Print pub_cadena
         
         frmdocu.Reportes.WindowTitle = frmdocu.Reportes.WindowTitle & " Archivo: " & Trim(frmdocu.Reportes.ReportFileName)
         frmdocu.Reportes.Action = 1
         On Error GoTo 0
       End If
PASA_OP:
       If LOC_TIPMOV = 10 Or LOC_TIPMOV = 100 Or LOC_TIPMOV = 3 Then
         WSNUMDOC = Left(frmdocu.Reportes.WindowTitle, 23)
         If LK_EMP = "PAR" Then
           WSNUMDOC = Right(Trim(frmdocu.Reportes.WindowTitle), 15)
         End If
         If LK_EMP = "3AA" Then
              WSNUMDOC = Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000") & ""
         End If
         frmdocu.Reportes.WindowTitle = "GUIA DE VENTA  " & Trim(frmdocu.Reportes.WindowTitle)
         If sin_valor.Value = 1 Then
            frmdocu.Reportes.ReportFileName = wRuta + "FACGUIA2.RPT"
         Else
            frmdocu.Reportes.ReportFileName = wRuta + Trim(pac_llave!PAC_ARCHI_GUIA) ' "FACGUIA.RPT"
         End If
         If LOC_TIPMOV = 100 Then
            If sin_valor.Value = 1 Then
            frmdocu.Reportes.ReportFileName = wRuta + "GUIAR2.RPT"
            Else
            frmdocu.Reportes.ReportFileName = wRuta + "GRE020.RPT"
         End If
         End If
         
         pub_mensaje = "Desea Imprimir la " & Trim(frmdocu.Reportes.WindowTitle) & "   ¿ Continuar... ?"
         Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
         If Pub_Respuesta = vbYes Then
           frmdocu.Reportes.SelectionFormula = pub_cadena
           Reportes.Formulas(21) = "FECHA_EMI=  '" & d_fecha_compra.Caption & "'"
           Reportes.Formulas(10) = ""
           Reportes.Formulas(11) = "FECHA_PARTIDA=  '" & FECHA_PART.Text & "'"
           Reportes.Formulas(12) = ""
           Reportes.Formulas(13) = ""
           Reportes.Formulas(14) = ""
           Reportes.Formulas(15) = ""
           Reportes.Formulas(16) = ""
           If cherela.Value = 1 And LOC_TIPMOV = 10 Then
             Reportes.Formulas(10) = "NUMDOC=  '" & WSNUMDOC & "'"
           End If
           Reportes.Formulas(22) = ""
           Reportes.Formulas(23) = ""
           Reportes.Formulas(24) = ""
           Reportes.Formulas(25) = ""
           Reportes.Formulas(26) = ""
           Reportes.Formulas(27) = ""
           Reportes.Formulas(28) = ""
            Select Case val(Right(cmbMotivo.Text, 6))
                 Case 1
                       Reportes.Formulas(22) = "VENT1= 'X'"
                 Case 2
                       Reportes.Formulas(23) = "VENT2= 'X'"
                 Case 3
                       Reportes.Formulas(24) = "VENT3= 'X'"
                 Case 4
                       Reportes.Formulas(25) = "VENT4= 'X'"
                 Case 5
                       Reportes.Formulas(26) = "VENT5= 'X'"
                 Case 6
                       Reportes.Formulas(27) = "VENT6= 'X'"
                 Case 7
                       Reportes.Formulas(28) = "VENT7= 'X'"
                 Case 8
                       Reportes.Formulas(26) = "VENT8= 'X'"
                 Case 9
                       Reportes.Formulas(27) = "VENT9= 'X'"
                 Case 10
                       Reportes.Formulas(28) = "VENT10= 'X'"
            End Select

           
           If chetrans.Value = 1 Then
             PS_TRA(0) = val(Right(TRANS.Text, 3))
             llave_trans.Requery
             Reportes.Formulas(12) = "TRN_NOMBRE    =  '" & llave_trans!TRN_NOMBRE & "'"
             Reportes.Formulas(13) = "TRN_DIRECCION =  '" & llave_trans!TRN_DIRECCION & "'"
             Reportes.Formulas(14) = "TRN_RUC       =  '" & llave_trans!TRN_RUC & "'"
             Reportes.Formulas(15) = "TRN_DNI       =  '" & llave_trans!TRN_DNI & "'"
             Reportes.Formulas(16) = "TRN_PLACA     =  '" & llave_trans!TRN_PLACA & "'"
             If LK_EMP = "PIU" Or LK_EMP = "3AA" Then
              Reportes.Formulas(17) = "TRN_CHOFER     =  '" & llave_trans!TRN_CHOFER & "'"
              Reportes.Formulas(18) = "TRN_DIR_CHOFER =  '" & llave_trans!TRN_DIR_CHOFER & "'"
              Reportes.Formulas(19) = "TRN_BREVETE    =  '" & llave_trans!TRN_BREVETE & "'"
              Reportes.Formulas(20) = "TRN_DNI_CHOFER =  '" & llave_trans!TRN_DNI_CHOFER & "'"
             End If
           End If
           frmdocu.Reportes.WindowTitle = frmdocu.Reportes.WindowTitle & " Archivo: " & Trim(frmdocu.Reportes.ReportFileName)
           'Debug.Print frmdocu.Reportes.WindowTitle
           On Error GoTo accion
           frmdocu.Reportes.Action = 1
           On Error GoTo 0
         End If
       End If
    Else
pasa_todo:
       frmdocu.Reportes.SelectionFormula = pub_cadena
       frmdocu.Reportes.WindowTitle = frmdocu.Reportes.WindowTitle & " Archivo: " & Trim(frmdocu.Reportes.ReportFileName)
       On Error GoTo accion
       frmdocu.Reportes.Action = 1
       On Error GoTo 0
       If cherela.Value = 1 Then
          GoTo PASA_OP
       End If
    End If
Exit Function
accion:
'Debug.Print pub_cadena
 MsgBox Err.Description
 MsgBox "Intente Nuevamente, la impresion de Modo manual", 48, Pub_Titulo
  Exit Function
Resume Next
End Function

Private Sub TxtSubZonaTrabajo_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  FRADIRE.Visible = False
  Exit Sub
End If

If KeyAscii = 13 Then
 SQ_OPER = 1
 pu_cp = "C"
 pu_codcia = LK_CODCIA
 pu_codclie = val(d_Codclie.Caption)
 LEER_CLI_LLAVE
 If cli_llave.EOF Then
   MsgBox "Cleinte no Existe NO Procede... ", 48, Pub_Titulo
   Exit Sub
 End If
 cli_llave.Edit
 cli_llave!CLI_TRAB_DIREC = Trim(txtdire.Text)
 cli_llave!CLI_TRAB_NUM = val(txtnum.Text)
 cli_llave!cli_TRAB_ZONA = val(Right(TxtZonaTrabajo.Text, 4))
 cli_llave!cli_TRAB_SUBZONA = val(Right(TxtSubZonaTrabajo.Text, 4))
 cli_llave.Update
 FRADIRE.Visible = False
 txtnumfac_KeyPress 13
End If



End Sub

Private Sub txtvend_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  If val(txtvend.Text) <> 0 Then
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    PUB_CODVEN = Nulo_Valor0(txtvend.Text)
    LEER_VEN_LLAVE
    If ven_llave.EOF Then
      MsgBox "Verificar Vendedor.", 48, Pub_Titulo
      Exit Sub
    End If
    cmbFBG.SetFocus
    Else
        cmbFBG.SetFocus
        Exit Sub
    End If
End If

End Sub
Public Sub ASIGNA_INT(WCONTROL As ComboBox, txt As Integer)
For fila = 0 To WCONTROL.ListCount - 1
    If val(Trim(Right(WCONTROL.List(fila), 3))) = txt Then
        WCONTROL.ListIndex = fila
        Exit Sub
    End If
Next fila
End Sub

Public Sub LLENA_ZONA(cont As ComboBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    PUB_CODCIA = "00"
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!tab_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub

Private Sub TxtZonaTrabajo_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  FRADIRE.Visible = False
  Exit Sub
End If
If KeyAscii = 13 Then
 TxtSubZonaTrabajo.SetFocus
  SendKeys "%{UP}"
End If


End Sub
