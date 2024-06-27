VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F6E4F630-E903-11D5-8BB9-0080AD40A177}#1.18#0"; "OSControlsUser.ocx"
Begin VB.Form frmPrecios 
   Caption         =   "Actualizacion de Precios"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmPrecios.frx":0000
   ScaleHeight     =   6855
   ScaleWidth      =   11130
   WindowState     =   2  'Maximized
   Begin VB.Frame frmParametros 
      Caption         =   "Definción de Parámetros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   6270
      Left            =   75
      TabIndex        =   31
      Top             =   810
      Visible         =   0   'False
      Width           =   3195
      Begin VB.CheckBox chkDesc4 
         BackColor       =   &H00808080&
         Caption         =   "Dscto.4(%)"
         Enabled         =   0   'False
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   420
         TabIndex        =   55
         Top             =   2550
         Width           =   1140
      End
      Begin VB.CheckBox chkDesc3 
         BackColor       =   &H00808080&
         Caption         =   "Dscto.3(%)"
         Enabled         =   0   'False
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   420
         TabIndex        =   54
         Top             =   2225
         Width           =   1140
      End
      Begin VB.CheckBox chkDesc2 
         BackColor       =   &H00808080&
         Caption         =   "Dscto.2(%)"
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   420
         TabIndex        =   53
         Top             =   1900
         Width           =   1140
      End
      Begin VB.CheckBox chkDesc1 
         BackColor       =   &H00808080&
         Caption         =   "Dscto.1(%)"
         Enabled         =   0   'False
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   420
         TabIndex        =   52
         Top             =   1575
         Value           =   1  'Checked
         Width           =   1140
      End
      Begin VB.CheckBox chkDescPre 
         BackColor       =   &H00808080&
         Caption         =   "Descuentos a Precios"
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
         Height          =   255
         Left            =   150
         TabIndex        =   49
         Top             =   3930
         Value           =   1  'Checked
         Width           =   2220
      End
      Begin OSControlsUser.ctlText txtFactorManejo 
         Height          =   285
         Left            =   1725
         TabIndex        =   11
         Top             =   825
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         BackColor       =   12632256
         ForeColor       =   -2147483642
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImputData       =   3
         Text            =   "0.000"
         LenDecimal      =   3
      End
      Begin VB.CheckBox chkDesc 
         BackColor       =   &H00808080&
         Caption         =   "Descuentos Sucesivos"
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
         Left            =   150
         TabIndex        =   12
         Top             =   1260
         Value           =   1  'Checked
         Width           =   2220
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Calculo con Prec. Unitario"
         Enabled         =   0   'False
         Height          =   210
         Index           =   1
         Left            =   105
         TabIndex        =   10
         Top             =   480
         Width           =   2370
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Calculo con CostRep de Lista"
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   210
         Value           =   -1  'True
         Width           =   2970
      End
      Begin OSControlsUser.ctlText txtPorDesc1 
         Height          =   285
         Left            =   1725
         TabIndex        =   13
         Top             =   1530
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         BackColor       =   12632256
         ForeColor       =   -2147483642
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImputData       =   3
         Text            =   "0.000"
         LenDecimal      =   3
      End
      Begin OSControlsUser.ctlText txtPorDesc2 
         Height          =   285
         Left            =   1725
         TabIndex        =   14
         Top             =   1860
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         BackColor       =   12632256
         ForeColor       =   -2147483642
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImputData       =   3
         Text            =   "0.000"
         LenDecimal      =   3
      End
      Begin OSControlsUser.ctlText txtPorDesc3 
         Height          =   285
         Left            =   1725
         TabIndex        =   15
         Top             =   2190
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         BackColor       =   12632256
         ForeColor       =   -2147483642
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImputData       =   3
         Text            =   "0.000"
         LenDecimal      =   3
      End
      Begin OSControlsUser.ctlText txtPorDesc4 
         Height          =   285
         Left            =   1725
         TabIndex        =   16
         Top             =   2520
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         BackColor       =   12632256
         ForeColor       =   -2147483642
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImputData       =   3
         Text            =   "0.000"
         LenDecimal      =   3
      End
      Begin OSControlsUser.ctlText txtUtilidad1 
         Height          =   285
         Left            =   1725
         TabIndex        =   19
         Top             =   4230
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         BackColor       =   12632256
         ForeColor       =   -2147483642
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImputData       =   3
         Text            =   "0.000"
         LenDecimal      =   3
      End
      Begin OSControlsUser.ctlText txtUtilidad2 
         Height          =   285
         Left            =   1725
         TabIndex        =   20
         Top             =   4560
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         BackColor       =   12632256
         ForeColor       =   -2147483642
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImputData       =   3
         Text            =   "0.000"
         LenDecimal      =   3
      End
      Begin OSControlsUser.ctlText txtUtilidad3 
         Height          =   285
         Left            =   1725
         TabIndex        =   21
         Top             =   4890
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         BackColor       =   12632256
         ForeColor       =   -2147483642
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImputData       =   3
         Text            =   "0.000"
         LenDecimal      =   3
      End
      Begin OSControlsUser.ctlText txtUtilidad4 
         Height          =   285
         Left            =   1725
         TabIndex        =   22
         Top             =   5220
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         BackColor       =   12632256
         ForeColor       =   -2147483642
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImputData       =   3
         Text            =   "0.000"
         LenDecimal      =   3
      End
      Begin OSControlsUser.ctlText txtUtilidad5 
         Height          =   285
         Left            =   1725
         TabIndex        =   23
         Top             =   5550
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         BackColor       =   12632256
         ForeColor       =   -2147483642
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImputData       =   3
         Text            =   "0.000"
         LenDecimal      =   3
      End
      Begin OSControlsUser.ctlText txtUtilidad6 
         Height          =   285
         Left            =   1725
         TabIndex        =   24
         Top             =   5880
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         BackColor       =   12632256
         ForeColor       =   -2147483642
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImputData       =   3
         Text            =   "0.000"
         LenDecimal      =   3
      End
      Begin OSControlsUser.ctlText txtPorAume1 
         Height          =   285
         Left            =   1725
         TabIndex        =   17
         Top             =   2955
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         BackColor       =   12632256
         ForeColor       =   -2147483642
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImputData       =   3
         Text            =   "0.000"
         LenDecimal      =   3
      End
      Begin OSControlsUser.ctlText txtPorAume2 
         Height          =   285
         Left            =   1725
         TabIndex        =   18
         Top             =   3285
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         BackColor       =   12632256
         ForeColor       =   -2147483642
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImputData       =   3
         Text            =   "0.000"
         LenDecimal      =   3
      End
      Begin VB.Label lblFactor 
         Caption         =   "Label2"
         Height          =   225
         Left            =   2400
         TabIndex        =   48
         Top             =   1275
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gasto.1(%)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   195
         Index           =   13
         Left            =   735
         TabIndex        =   46
         Top             =   2970
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gasto.2(%)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   195
         Index           =   12
         Left            =   735
         TabIndex        =   45
         Top             =   3300
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(%)Descuento 6"
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
         Height          =   285
         Index           =   11
         Left            =   420
         TabIndex        =   38
         Top             =   5925
         Width           =   1230
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(%)Descuento 5"
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
         Height          =   285
         Index           =   10
         Left            =   420
         TabIndex        =   37
         Top             =   5595
         Width           =   1230
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(%)Descuento 4"
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
         Height          =   285
         Index           =   8
         Left            =   420
         TabIndex        =   36
         Top             =   5265
         Width           =   1230
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(%)Descuento 3"
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
         Height          =   285
         Index           =   2
         Left            =   420
         TabIndex        =   35
         Top             =   4935
         Width           =   1230
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(%)Descuento 2"
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
         Height          =   285
         Index           =   6
         Left            =   420
         TabIndex        =   34
         Top             =   4605
         Width           =   1230
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(%)Descuento 1"
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
         Height          =   285
         Index           =   7
         Left            =   420
         TabIndex        =   33
         Top             =   4275
         Width           =   1230
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Factor Manejo"
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
         Left            =   570
         TabIndex        =   32
         Top             =   870
         Width           =   1035
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
         Height          =   1665
         Index           =   5
         Left            =   30
         TabIndex        =   44
         Top             =   1185
         Width           =   3120
      End
      Begin VB.Label lblart 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00404000&
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
         Height          =   750
         Index           =   6
         Left            =   45
         TabIndex        =   47
         Top             =   2880
         Width           =   3120
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
         Height          =   2295
         Index           =   7
         Left            =   30
         TabIndex        =   50
         Top             =   3930
         Width           =   3120
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1515
      Index           =   1
      Left            =   15
      TabIndex        =   0
      Top             =   -75
      Width           =   11760
      Begin VB.ComboBox cbo_Unidad 
         BackColor       =   &H00808080&
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmPrecios.frx":0342
         Left            =   3030
         List            =   "frmPrecios.frx":034C
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "S"
         Top             =   1095
         Width           =   1515
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   330
         Left            =   9948
         TabIndex        =   26
         Top             =   1080
         Width           =   915
      End
      Begin VB.CommandButton cmdcerrar 
         Caption         =   "Cer&rar"
         Height          =   330
         Left            =   10950
         TabIndex        =   27
         Top             =   1080
         Width           =   720
      End
      Begin VB.CommandButton cmdParametros 
         Caption         =   "Parametros"
         Height          =   330
         Left            =   7376
         TabIndex        =   8
         Top             =   1080
         Width           =   1200
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
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   135
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   435
         Width           =   2715
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
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   3045
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   435
         Width           =   2715
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
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   5970
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   435
         Width           =   2715
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
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   8880
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   435
         Width           =   2715
      End
      Begin VB.CommandButton cmdconsultar 
         Caption         =   "Consultar"
         Height          =   330
         Left            =   6090
         TabIndex        =   7
         Top             =   1080
         Width           =   1200
      End
      Begin VB.CommandButton cmdCalcular 
         Caption         =   "Calcular"
         Enabled         =   0   'False
         Height          =   330
         Left            =   8662
         TabIndex        =   25
         Top             =   1080
         Width           =   1200
      End
      Begin VB.ComboBox cbo_ViewMoney 
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmPrecios.frx":035F
         Left            =   165
         List            =   "frmPrecios.frx":0369
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "S"
         Top             =   1110
         Width           =   2475
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad :"
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
         Left            =   3045
         TabIndex        =   51
         Top             =   885
         Width           =   600
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
         Left            =   165
         TabIndex        =   42
         Top             =   180
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
         Left            =   3090
         TabIndex        =   41
         Top             =   195
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
         Left            =   6030
         TabIndex        =   40
         Top             =   165
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00F8DED7&
         Height          =   195
         Index           =   18
         Left            =   8895
         TabIndex        =   39
         Top             =   180
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Precios en :"
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
         Index           =   14
         Left            =   180
         TabIndex        =   28
         Top             =   900
         Width           =   840
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
         Height          =   735
         Index           =   4
         Left            =   30
         TabIndex        =   43
         Top             =   135
         Width           =   11685
      End
   End
   Begin OSControlsUser.OSNewFlexGrid grid_unid 
      Height          =   5460
      Left            =   135
      TabIndex        =   29
      Top             =   1605
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   9631
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      BorderStyle     =   0
      BackColorSel    =   8388608
      ColEdit(1)      =   0   'False
      ColSalto(1)     =   0
      ColAtras(1)     =   0
      ColWidth0       =   960
   End
   Begin ComctlLib.ProgressBar PROGRESO 
      Height          =   210
      Left            =   165
      TabIndex        =   30
      Top             =   7065
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   370
      _Version        =   327682
      Appearance      =   0
   End
End
Attribute VB_Name = "frmPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RDQPRECIOS As rdoQuery
Dim RDRPRECIOS As rdoResultset

Dim Flag_Inicial As String
Dim FLAG As Integer
Dim sIndex As Integer
Dim FlagPrecio As Integer
Dim FactorDesc As Double
Dim FactorAume As Double
Dim FactorManejo As Double

Dim DescSuc As Integer
Dim DescPre As Integer
Const ColPre1 = 21
Const ColPre2 = 23
Const ColPre3 = 25
Const ColPre4 = 27
Const ColPre5 = 29
Const ColPre6 = 31

Dim iUtilidad1 As Double
Dim iUtilidad2 As Double
Dim iUtilidad3 As Double
Dim iUtilidad4 As Double
Dim iUtilidad5 As Double
Dim iUtilidad6 As Double
Dim vPor1 As Double
Dim vPor2 As Double
Dim vPor3 As Double
Dim vPor4 As Double
Dim dFactorManejo As Double
Dim dMargen As Double
Dim dLPP As Double
Dim WSPOR  As Double

Private Sub art_grupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        art_numero.SetFocus
        DoEvents
        SendKeys "%{UP}"
        Exit Sub
    End If
End Sub

Private Sub art_numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdconsultar.SetFocus
    End If
End Sub

Private Sub art_subfam_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        art_grupo.SetFocus
        DoEvents
        SendKeys "%{UP}"
        Exit Sub
    End If
End Sub

Private Sub art_subfam_LostFocus()
Dim wpos As Integer
Dim WFAMI2 As Integer
    If Trim(art_subfam.Text) = "" Then
        art_grupo.Clear
        Exit Sub
    End If
    wpos = art_grupo.ListIndex
    WFAMI2 = Val(Trim(Right(art_subfam.Text, 6)))
    PUB_TIPREG = 129
    LLENADO_SUBFAM art_grupo, WFAMI2
    On Error GoTo sigue
    art_grupo.ListIndex = wpos
    Exit Sub
sigue:
    Resume Next
End Sub

Private Sub chkDesc2_Click()
    If chkDesc2.Value = 1 Then
        chkDesc3.Enabled = True
        txtPorDesc2.Enabled = True
        grid_unid.ColWidth(9) = 800  'PorDesc2
        grid_unid.ColControl(9) = osTextBox 'Porc de Descto 1
        grid_unid.ColEdit(9) = True
        grid_unid.ColData(9) = osDecimal
        grid_unid.ColSalto(9) = 8
        CalFactorDesc
        txtPorDesc2.SetFocus
    Else
        If chkDesc3.Value = 0 Then
            chkDesc3.Enabled = False
            txtPorDesc2.Text = "0.000"
            txtPorDesc2.Enabled = False
            grid_unid.ColWidth(9) = 0  'PorDesc2
            Call LLenaColVal(0#, 9)
            grid_unid.ColEdit(9) = False
            CalFactorDesc
        Else
            chkDesc2.Value = 1
        End If
    End If
End Sub

Private Sub chkDesc3_Click()
    If chkDesc3.Value = 1 Then
        chkDesc4.Enabled = True
        txtPorDesc3.Enabled = True
        grid_unid.ColWidth(10) = 800  'PorDesc3
        grid_unid.ColEdit(10) = True
        grid_unid.ColData(10) = osDecimal
        grid_unid.ColSalto(10) = 7
        CalFactorDesc
        txtPorDesc3.SetFocus
    Else
        If chkDesc4.Value = 0 Then
            chkDesc4.Enabled = False
            txtPorDesc3.Text = "0.000"
            txtPorDesc3.Enabled = False
            grid_unid.ColWidth(10) = 0  'PorDesc3
            Call LLenaColVal(0#, 10)
            grid_unid.ColEdit(9) = False
            CalFactorDesc
        Else
            chkDesc3.Value = 1
        End If
    End If
End Sub

Private Sub chkDesc4_Click()
    If chkDesc4.Value = 1 Then
        grid_unid.ColWidth(11) = 800  'PorDesc4
        txtPorDesc4.Enabled = True
        grid_unid.ColEdit(11) = True
        grid_unid.ColData(11) = osDecimal
        grid_unid.ColSalto(11) = 6
        txtPorDesc4.SetFocus
    Else
        grid_unid.ColWidth(11) = 0  'PorDesc4
        txtPorDesc4.Text = "0.000"
        txtPorDesc4.Enabled = False
        Call LLenaColVal(0#, 11)
        grid_unid.ColEdit(9) = False
    End If
    CalFactorDesc
End Sub

Private Sub cmdCalcular_Click()
Dim NroRows  As Integer
Dim iRow As Long

    FlagPrecio = 2
    dLPP = 0
    dMargen = 0
    NroRows = grid_unid.Rows - 1
    For iRow = 1 To NroRows
        dMargen = dDouble(grid_unid.TextMatrix(iRow, 17))
        dLPP = dDouble(grid_unid.TextMatrix(iRow, 6))
        grid_unid.TextMatrix(iRow, ColPre1 - 1) = Val(txtUtilidad1.Text)
        grid_unid.TextMatrix(iRow, ColPre2 - 1) = Val(txtUtilidad2.Text)
        grid_unid.TextMatrix(iRow, ColPre3 - 1) = Val(txtUtilidad3.Text)
        grid_unid.TextMatrix(iRow, ColPre4 - 1) = Val(txtUtilidad4.Text)
        grid_unid.TextMatrix(iRow, ColPre5 - 1) = Val(txtUtilidad5.Text)
        grid_unid.TextMatrix(iRow, ColPre6 - 1) = Val(txtUtilidad6.Text)
        
        Call LLenaColVal(txtPorDesc1.Text, 8)
        Call LLenaColVal(txtPorDesc2.Text, 9)
        Call LLenaColVal(txtPorDesc3.Text, 10)
        Call LLenaColVal(txtPorDesc4.Text, 11)
        
        CalculaLinea iRow, dLPP, dMargen
    Next iRow
End Sub
Private Sub CalculaLinea(ByVal iRow As Long, ByVal LPP As Double, ByVal Margen As Double)
Dim iPrecio1 As Double
Dim iPrecio2 As Double
Dim iPrecio3 As Double
Dim iPrecio4 As Double
Dim iPrecio5 As Double
Dim iPrecio6 As Double
Dim iFactorDescxLin As Double
Dim CVV As Double
Dim RM As Double
On Error GoTo Handler
        WSPOR = 0
        vPor1 = dDouble(grid_unid.TextMatrix(iRow, 8))
        vPor2 = dDouble(grid_unid.TextMatrix(iRow, 9))
        vPor3 = dDouble(grid_unid.TextMatrix(iRow, 10))
        vPor4 = dDouble(grid_unid.TextMatrix(iRow, 11))
        
        iUtilidad1 = Val(txtUtilidad1.Text)
        iUtilidad2 = Val(txtUtilidad2.Text)
        iUtilidad3 = Val(txtUtilidad3.Text)
        iUtilidad4 = Val(txtUtilidad4.Text)
        iUtilidad5 = Val(txtUtilidad5.Text)
        iUtilidad6 = Val(txtUtilidad6.Text)
        
        If DescSuc = 1 Then
            iFactorDescxLin = 100 - (100 * ((100 - vPor1) / 100)) * ((100 - vPor2) / 100) * ((100 - vPor3) / 100) * ((100 - vPor4) / 100)
        Else
            iFactorDescxLin = 0
        End If
        
        grid_unid.TextMatrix(iRow, 13) = iFactorDescxLin
        
        LPP = grid_unid.TextMatrix(iRow, 6)
        CVV = LPP - (LPP * iFactorDescxLin / 100)
        grid_unid.TextMatrix(iRow, 15) = Format(CVV, "0.000")
        grid_unid.TextMatrix(iRow, 16) = Format(CVV * (1 + LK_IGV / 100), "0.000")
        RM = CVV / ((100 - Margen) / 100)
        grid_unid.TextMatrix(iRow, 18) = Format(RM, "0.000")
        grid_unid.TextMatrix(iRow, 19) = Format(RM / ((100 - dFactorManejo) / 100), "0.000")
        
        If DescPre = 1 Then
            If FlagPrecio = 1 Then
                If Not RDRPRECIOS.EOF Then
                    iPrecio1 = RDRPRECIOS("PRE_PRE1")
                    iPrecio2 = RDRPRECIOS("PRE_PRE2")
                    iPrecio3 = RDRPRECIOS("PRE_PRE3")
                    iPrecio4 = RDRPRECIOS("PRE_PRE4")
                    iPrecio5 = RDRPRECIOS("PRE_PRE5")
                    iPrecio6 = RDRPRECIOS("PRE_PRE6")
                End If
             Else
                iPrecio1 = (grid_unid.TextMatrix(iRow, 19) - (grid_unid.TextMatrix(iRow, 19) * (iUtilidad1 / 100))) * (1 + LK_IGV / 100)
                iPrecio2 = (grid_unid.TextMatrix(iRow, 19) - (grid_unid.TextMatrix(iRow, 19) * (iUtilidad2 / 100))) * (1 + LK_IGV / 100)
                iPrecio3 = (grid_unid.TextMatrix(iRow, 19) - (grid_unid.TextMatrix(iRow, 19) * (iUtilidad3 / 100))) * (1 + LK_IGV / 100)
                iPrecio4 = (grid_unid.TextMatrix(iRow, 19) - (grid_unid.TextMatrix(iRow, 19) * (iUtilidad4 / 100))) * (1 + LK_IGV / 100)
                iPrecio5 = (grid_unid.TextMatrix(iRow, 19) - (grid_unid.TextMatrix(iRow, 19) * (iUtilidad5 / 100))) * (1 + LK_IGV / 100)
                iPrecio6 = (grid_unid.TextMatrix(iRow, 19) - (grid_unid.TextMatrix(iRow, 19) * (iUtilidad6 / 100))) * (1 + LK_IGV / 100)
            End If
            
        Else
'            iPrecio1 = iPrecio * (1 + iUtilidad1)
'            iPrecio2 = iPrecio * (1 + iUtilidad2)
'            iPrecio3 = iPrecio * (1 + iUtilidad3)
'            iPrecio4 = iPrecio * (1 + iUtilidad4)
'            iPrecio5 = iPrecio * (1 + iUtilidad5)
'            iPrecio6 = iPrecio * (1 + iUtilidad6)
        End If
        
        grid_unid.TextMatrix(iRow, ColPre1) = Format(iPrecio1, "0.000")
        'If Val(grid_unid.TextMatrix(iRow, 6)) <> 0 Then WSPOR = (Nulo_Valor0(iPrecio1) * 100) / Val(grid_unid.TextMatrix(iRow, 6)) - 100
        'grid_unid.TextMatrix(iRow, ColPre1 - 1) = Format(WSPOR, "0.00")
        If cbo_ViewMoney.ListIndex = 1 Then '' .Tag = "D" Then
          grid_unid.TextMatrix(iRow, ColPre1 + 11) = Format(iPrecio1, "0.000")
        Else
          grid_unid.TextMatrix(iRow, ColPre1 + 16) = Format(iPrecio1, "0.000")
        End If
        If FlagPrecio = 1 And dDouble(grid_unid.TextMatrix(iRow, ColPre1)) <> 0 Then
            CALCULAR_POR1 grid_unid.TextMatrix(iRow, ColPre1), ColPre1, iRow
        Else
            grid_unid.TextMatrix(iRow, ColPre1 - 1) = Format(iUtilidad1, "0.000")
        End If
        
        grid_unid.TextMatrix(iRow, ColPre2) = Format(iPrecio2, "0.000")
        'If Val(grid_unid.TextMatrix(iRow, 6)) <> 0 Then WSPOR = (Nulo_Valor0(iPrecio2) * 100) / Val(grid_unid.TextMatrix(iRow, 6)) - 100
        'grid_unid.TextMatrix(iRow, ColPre2 - 1) = Format(WSPOR, "0.00")
        If cbo_ViewMoney.ListIndex = 1 Then ''.Tag = "D" Then
          grid_unid.TextMatrix(iRow, ColPre2 + 11) = Format(iPrecio2, "0.000")
        Else
          grid_unid.TextMatrix(iRow, ColPre2 + 16) = Format(iPrecio2, "0.000")
        End If
        If FlagPrecio = 1 And dDouble(grid_unid.TextMatrix(iRow, ColPre2)) <> 0 Then
            CALCULAR_POR1 grid_unid.TextMatrix(iRow, ColPre2), ColPre2, iRow
        Else
            grid_unid.TextMatrix(iRow, ColPre2 - 1) = Format(iUtilidad2, "0.000")
        End If
        
        grid_unid.TextMatrix(iRow, ColPre3) = Format(iPrecio3, "0.000")
        'If Val(grid_unid.TextMatrix(iRow, 6)) <> 0 Then WSPOR = (Nulo_Valor0(iPrecio3) * 100) / Val(grid_unid.TextMatrix(iRow, 6)) - 100
        'grid_unid.TextMatrix(iRow, ColPre3 - 1) = Format(WSPOR, "0.00")
        If cbo_ViewMoney.ListIndex = 1 Then ''.Tag = "D" Then
          grid_unid.TextMatrix(iRow, ColPre3 + 11) = Format(iPrecio3, "0.000")
        Else
          grid_unid.TextMatrix(iRow, ColPre3 + 16) = Format(iPrecio3, "0.000")
        End If
        If FlagPrecio = 1 And dDouble(grid_unid.TextMatrix(iRow, ColPre3)) <> 0 Then
            CALCULAR_POR1 grid_unid.TextMatrix(iRow, ColPre3), ColPre3, iRow
        Else
            grid_unid.TextMatrix(iRow, ColPre3 - 1) = Format(iUtilidad3, "0.000")
        End If
        
        grid_unid.TextMatrix(iRow, ColPre4) = Format(iPrecio4, "0.000")
        'If Val(grid_unid.TextMatrix(iRow, 6)) <> 0 Then WSPOR = (Nulo_Valor0(iPrecio4) * 100) / Val(grid_unid.TextMatrix(iRow, 6)) - 100
        'grid_unid.TextMatrix(iRow, ColPre4 - 1) = Format(WSPOR, "0.00")
        If cbo_ViewMoney.ListIndex = 1 Then ''.Tag = "D" Then
          grid_unid.TextMatrix(iRow, ColPre4 + 11) = Format(iPrecio4, "0.000")
        Else
          grid_unid.TextMatrix(iRow, ColPre4 + 16) = Format(iPrecio4, "0.000")
        End If
        If FlagPrecio = 1 And dDouble(grid_unid.TextMatrix(iRow, ColPre4)) <> 0 Then
            CALCULAR_POR1 grid_unid.TextMatrix(iRow, ColPre4), ColPre4, iRow
        Else
            grid_unid.TextMatrix(iRow, ColPre4 - 1) = Format(iUtilidad4, "0.000")
        End If
        
        grid_unid.TextMatrix(iRow, ColPre5) = Format(iPrecio5, "0.000")
        'If Val(grid_unid.TextMatrix(iRow, 6)) <> 0 Then WSPOR = (Nulo_Valor0(iPrecio5) * 100) / Val(grid_unid.TextMatrix(iRow, 6)) - 100
        'grid_unid.TextMatrix(iRow, ColPre5 - 1) = Format(WSPOR, "0.00")
        If cbo_ViewMoney.ListIndex = 1 Then ''.Tag = "D" Then
          grid_unid.TextMatrix(iRow, ColPre5 + 11) = Format(iPrecio5, "0.000")
        Else
          grid_unid.TextMatrix(iRow, ColPre5 + 16) = Format(iPrecio5, "0.000")
        End If
        If FlagPrecio = 1 And dDouble(grid_unid.TextMatrix(iRow, ColPre5)) <> 0 Then
            CALCULAR_POR1 grid_unid.TextMatrix(iRow, ColPre5), ColPre5, iRow
        Else
            grid_unid.TextMatrix(iRow, ColPre5 - 1) = Format(iUtilidad5, "0.000")
        End If
        
        grid_unid.TextMatrix(iRow, ColPre6) = Format(iPrecio6, "0.000")
        'If Val(grid_unid.TextMatrix(iRow, 6)) <> 0 Then WSPOR = (Nulo_Valor0(iPrecio5) * 100) / Val(grid_unid.TextMatrix(iRow, 6)) - 100
        'grid_unid.TextMatrix(iRow, ColPre6 - 1) = Format(WSPOR, "0.00")
        If cbo_ViewMoney.ListIndex = 1 Then ''.Tag = "D" Then
          grid_unid.TextMatrix(iRow, ColPre6 + 11) = Format(iPrecio6, "0.000")
        Else
          grid_unid.TextMatrix(iRow, ColPre6 + 16) = Format(iPrecio6, "0.000")
        End If
        If FlagPrecio = 1 And dDouble(grid_unid.TextMatrix(iRow, ColPre6)) <> 0 Then
            CALCULAR_POR1 grid_unid.TextMatrix(iRow, ColPre6), ColPre6, iRow
        Else
            grid_unid.TextMatrix(iRow, ColPre6 - 1) = Format(iUtilidad6, "0.000")
        End If
        Exit Sub
Handler:
    MsgBox Err.Description
    Err.Clear
End Sub
Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdconsultar_Click()
Dim sCodigo As String
Dim SQL As String
Dim sMoneda As String
Dim UNIDAD As String
Dim Familia As Integer
Dim SubFam As Integer
Dim grupo As Integer
Dim NUMERO As Integer
On Error GoTo Handler
    FlagPrecio = 1
    
    Familia = Val(Right(art_familia.Text, 6))
    SubFam = Val(Right(art_subfam.Text, 6))
    grupo = Val(Right(art_grupo.Text, 6))
    NUMERO = Val(Right(art_numero.Text, 6))
    
    If cbo_Unidad.ListIndex = 0 Then
        UNIDAD = " AND PRE_FLAG_UNIDAD = 'A' "
    ElseIf cbo_Unidad.ListIndex = 1 Then
        UNIDAD = ""
    End If
    
    sMoneda = "S"
    cbo_ViewMoney.ListIndex = 0
        
    SQL = "SELECT ARTI.ART_NOMBRE AS Articulo, ARTI.ART_ALTERNO AS Alterno, PRECIOS.* "
    SQL = SQL & "FROM ARTI INNER JOIN PRECIOS ON ARTI.ART_KEY = PRECIOS.PRE_CODART AND ARTI.ART_CODCIA = PRECIOS.PRE_CODCIA "
    SQL = SQL & "WHERE ARTI.ART_CODCIA = '" & LK_CODCIA & "' " & UNIDAD
    SQL = SQL & " AND ART_FAMILIA = " & Familia
    SQL = SQL & " AND ART_SUBFAM = " & SubFam
    SQL = SQL & " AND ART_SUBGRU = " & grupo
    SQL = SQL & " AND ART_NUMERO = " & NUMERO
    Set RDQPRECIOS = CN.CreateQuery("", SQL)
    Set RDRPRECIOS = RDQPRECIOS.OpenResultset(rdOpenKeyset, rdConcurValues)
    RDRPRECIOS.Requery
    PROGRESO.Visible = True
    grid_unid.Clear
    
    grid_unid.Rows = IIf(RDRPRECIOS.RowCount = 0, 2, RDRPRECIOS.RowCount + 1)
    FLAG = IIf(RDRPRECIOS.RowCount = 0, 0, 1)
    If FLAG = 1 Then
        cmdCalcular.Enabled = True
        cbo_ViewMoney.Enabled = True
    Else
        cmdCalcular.Enabled = False
        cbo_ViewMoney.Enabled = False
    End If
    PROGRESO.max = grid_unid.Rows
    fila = 0
    grid_unid.Visible = False
    Do While Not RDRPRECIOS.EOF
        fila = fila + 1
        PROGRESO.Value = fila
        'grid_unid.RowHeight(fila) = 285
        grid_unid.Row = fila
        grid_unid.TextMatrix(fila, 0) = Trim(RDRPRECIOS!articulo)
        grid_unid.TextMatrix(fila, 1) = Trim(RDRPRECIOS!Alterno)
        grid_unid.TextMatrix(fila, 2) = Trim(RDRPRECIOS!pre_unidad)
        grid_unid.TextMatrix(fila, 3) = RDRPRECIOS!PRE_EQUIV
        
        grid_unid.TextMatrix(fila, 6) = RDRPRECIOS("PRE_LITRO")
        grid_unid.TextMatrix(fila, 48) = RDRPRECIOS("PRE_LITRO")
        grid_unid.TextMatrix(fila, 17) = RDRPRECIOS("PRE_COSTO_REPO")
        grid_unid.TextMatrix(fila, 8) = Nulo_Valor0(RDRPRECIOS!PRE_PORDES1)
        
        grid_unid.TextMatrix(fila, 9) = Nulo_Valor0(RDRPRECIOS!PRE_PORDES2)
        If Nulo_Valor0(RDRPRECIOS!PRE_PORDES2) <> 0 Then grid_unid.ColWidth(9) = 800
        grid_unid.TextMatrix(fila, 10) = Nulo_Valor0(RDRPRECIOS!PRE_PORDES3)
        If Nulo_Valor0(RDRPRECIOS!PRE_PORDES3) <> 0 Then grid_unid.ColWidth(10) = 800
        grid_unid.TextMatrix(fila, 11) = Nulo_Valor0(RDRPRECIOS!PRE_PORDES4)
        If Nulo_Valor0(RDRPRECIOS!PRE_PORDES4) <> 0 Then grid_unid.ColWidth(11) = 800
        
        CalculaLinea fila, RDRPRECIOS("PRE_LITRO"), RDRPRECIOS("PRE_COSTO_REPO")
        'Para precio 1 ===================================================
        grid_unid.COL = ColPre1 - 1
        grid_unid.CellForeColor = QBColor(9)
        If Val(grid_unid.TextMatrix(fila, 6)) <> 0 Then WSPOR = (Nulo_Valor0(RDRPRECIOS!PRE_PRE1) * 100) / Val(grid_unid.TextMatrix(fila, 6)) - 100
        'grid_unid.TextMatrix(fila, ColPre1 - 1) = Format(WSPOR, "0.00")
        If sMoneda = "S" Then
          grid_unid.TextMatrix(fila, ColPre1) = Nulo_Valor0(RDRPRECIOS!PRE_PRE1)
        Else
          grid_unid.TextMatrix(fila, ColPre1) = Nulo_Valor0(RDRPRECIOS!pre_pre11)
        End If
        grid_unid.TextMatrix(fila, ColPre1 + 11) = Nulo_Valor0(RDRPRECIOS!pre_pre11)
        grid_unid.TextMatrix(fila, ColPre1 + 16) = Nulo_Valor0(RDRPRECIOS!PRE_PRE1)
        'Para precio 2 ===================================================
        grid_unid.COL = ColPre2 - 1
        grid_unid.CellForeColor = QBColor(9)
        If Val(grid_unid.TextMatrix(fila, 6)) <> 0 Then WSPOR = (Nulo_Valor0(RDRPRECIOS!PRE_PRE2) * 100) / Val(grid_unid.TextMatrix(fila, 6)) - 100
        'grid_unid.TextMatrix(fila, ColPre2 - 1) = Format(WSPOR, "0.00")
        If sMoneda = "S" Then
         grid_unid.TextMatrix(fila, ColPre2) = Nulo_Valor0(RDRPRECIOS!PRE_PRE2)
        Else
         grid_unid.TextMatrix(fila, ColPre2) = Nulo_Valor0(RDRPRECIOS!PRE_PRE22)
        End If
        grid_unid.TextMatrix(fila, ColPre2 + 11) = Nulo_Valor0(RDRPRECIOS!PRE_PRE22)
        grid_unid.TextMatrix(fila, ColPre2 + 16) = Nulo_Valor0(RDRPRECIOS!PRE_PRE2)
        'Para precio 3 ===================================================
        grid_unid.COL = ColPre3 - 1
        grid_unid.CellForeColor = QBColor(9)
        If Val(grid_unid.TextMatrix(fila, 6)) <> 0 Then WSPOR = (Nulo_Valor0(RDRPRECIOS!PRE_PRE3) * 100) / Val(grid_unid.TextMatrix(fila, 6)) - 100
        'grid_unid.TextMatrix(fila, ColPre3 - 1) = Format(WSPOR, "0.00")
        If sMoneda = "S" Then
         grid_unid.TextMatrix(fila, ColPre3) = Nulo_Valor0(RDRPRECIOS!PRE_PRE3)
        Else
         grid_unid.TextMatrix(fila, ColPre3) = Nulo_Valor0(RDRPRECIOS!PRE_PRE33)
        End If
        grid_unid.TextMatrix(fila, ColPre3 + 11) = Nulo_Valor0(RDRPRECIOS!PRE_PRE33)
        grid_unid.TextMatrix(fila, ColPre3 + 16) = Nulo_Valor0(RDRPRECIOS!PRE_PRE3)
        'Para precio 4 ===================================================
        grid_unid.COL = ColPre4 - 1
        grid_unid.CellForeColor = QBColor(9)
        If Val(grid_unid.TextMatrix(fila, 6)) <> 0 Then WSPOR = (Nulo_Valor0(RDRPRECIOS!PRE_PRE4) * 100) / Val(grid_unid.TextMatrix(fila, 6)) - 100
        'grid_unid.TextMatrix(fila, ColPre4 - 1) = Format(WSPOR, "0.00")
        If sMoneda = "S" Then
          grid_unid.TextMatrix(fila, ColPre4) = Nulo_Valor0(RDRPRECIOS!PRE_PRE4)
        Else
          grid_unid.TextMatrix(fila, ColPre4) = Nulo_Valor0(RDRPRECIOS!PRE_PRE44)
        End If
        grid_unid.TextMatrix(fila, ColPre4 + 11) = Nulo_Valor0(RDRPRECIOS!PRE_PRE44)
        grid_unid.TextMatrix(fila, ColPre4 + 16) = Nulo_Valor0(RDRPRECIOS!PRE_PRE4)
        'Para precio 5 ===================================================
        grid_unid.COL = ColPre5 - 1
        grid_unid.CellForeColor = QBColor(9)
        If Val(grid_unid.TextMatrix(fila, 6)) <> 0 Then WSPOR = (Nulo_Valor0(RDRPRECIOS!PRE_PRE5) * 100) / Val(grid_unid.TextMatrix(fila, 6)) - 100
        'grid_unid.TextMatrix(fila, ColPre5 - 1) = Format(WSPOR, "0.00")
        If sMoneda = "S" Then
         grid_unid.TextMatrix(fila, ColPre5) = Nulo_Valor0(RDRPRECIOS!PRE_PRE5)
        Else
         grid_unid.TextMatrix(fila, ColPre5) = Nulo_Valor0(RDRPRECIOS!PRE_PRE55)
        End If
        grid_unid.TextMatrix(fila, ColPre5 + 11) = Nulo_Valor0(RDRPRECIOS!PRE_PRE55)
        grid_unid.TextMatrix(fila, ColPre5 + 16) = Nulo_Valor0(RDRPRECIOS!PRE_PRE5)
        'Para precio 6 ===================================================
        grid_unid.COL = ColPre6 - 1
        grid_unid.CellForeColor = QBColor(9)
        If Val(grid_unid.TextMatrix(fila, 6)) <> 0 Then WSPOR = (Nulo_Valor0(RDRPRECIOS!PRE_PRE6) * 100) / Val(grid_unid.TextMatrix(fila, 6)) - 100
        'grid_unid.TextMatrix(fila, ColPre6 - 1) = Format(WSPOR, "0.00")
        If sMoneda = "S" Then
         grid_unid.TextMatrix(fila, ColPre6) = Nulo_Valor0(RDRPRECIOS!PRE_PRE6)
        Else
         grid_unid.TextMatrix(fila, ColPre6) = Nulo_Valor0(RDRPRECIOS!PRE_PRE66)
        End If
        grid_unid.TextMatrix(fila, ColPre6 + 11) = Nulo_Valor0(RDRPRECIOS!PRE_PRE66)
        grid_unid.TextMatrix(fila, ColPre6 + 16) = Nulo_Valor0(RDRPRECIOS!PRE_PRE6)
        'Para otros ===================================================
        grid_unid.TextMatrix(fila, 46) = Nulo_Valor0(RDRPRECIOS!PRE_codart)
        grid_unid.TextMatrix(fila, 47) = Nulo_Valor0(RDRPRECIOS!pre_secuencia)
        
        RDRPRECIOS.MoveNext
    Loop
    grid_unid.Visible = True
    PROGRESO.Value = 1
    PROGRESO.Visible = False
    grid_unid.COL = 1
    grid_unid.Row = 1
'    Option2(0).SetFocus
    Exit Sub
Handler:
    PROGRESO.Value = 1
    PROGRESO.Visible = False
    MsgBox Err.Description, vbCritical, Pub_Titulo
End Sub


Private Sub art_familia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        art_subfam.SetFocus
        DoEvents
        SendKeys "%{UP}"
        Exit Sub
    End If
End Sub

Private Sub art_familia_LostFocus()
Dim wpos As Integer
Dim WFAMI2 As Integer
    If Trim(art_familia.Text) = "" Then
        art_subfam.Clear
        Exit Sub
    End If
    wpos = art_subfam.ListIndex
    WFAMI2 = Val(Trim(Right(art_familia.Text, 6)))
    PUB_TIPREG = 123
    LLENADO_SUBFAM art_subfam, WFAMI2
    On Error GoTo sigue
    art_subfam.ListIndex = wpos
    Exit Sub
sigue:
    Resume Next
End Sub

Private Sub CalFactorDesc()

    vPor1 = Val(txtPorDesc1.Text)
    vPor2 = Val(txtPorDesc2.Text)
    vPor3 = Val(txtPorDesc3.Text)
    vPor4 = Val(txtPorDesc4.Text)
    
    If DescSuc = 1 Then
        FactorDesc = 100 - (100 * ((100 - vPor1) / 100)) * ((100 - vPor2) / 100) * ((100 - vPor3) / 100) * ((100 - vPor4) / 100)
    Else
        FactorDesc = 0
    End If
    lblFactor.Caption = Format(FactorDesc, "0.00")
    LLenaColVal FactorDesc, 13
End Sub
Private Sub CalFactorAume()
Dim vPor1 As Double
Dim vPor2 As Double

    vPor1 = Val(txtPorDesc1.Text)
    vPor2 = Val(txtPorDesc2.Text)
    
    If DescSuc = 1 Then
        FactorAume = 100 - (100 * ((100 - vPor1) / 100)) * ((100 - vPor2) / 100)
    Else
        FactorAume = 0
    End If
'    lblFactor.Caption = Format(Factor, "0.00")
End Sub
Public Sub CALCULAR_POR(WSPRE As Currency, WSCOL As Long)
Dim valor As Currency
    If Val(grid_unid.TextMatrix(grid_unid.Row, 6)) <> 0 Then
      valor = (WSPRE * 100) / Val(grid_unid.TextMatrix(grid_unid.Row, 6)) - 100
    Else
      valor = 0
    End If
    grid_unid.COL = WSCOL
    grid_unid.TextMatrix(grid_unid.Row, WSCOL - 1) = Format(valor, "0.00")
End Sub
Public Sub CALCULAR_PRE(WSPOR As Currency, WSCOL As Long)
Dim valor As Currency

    valor = Val(grid_unid.TextMatrix(grid_unid.Row, 6)) * (1 + (WSPOR / 100))
    grid_unid.COL = WSCOL
    grid_unid.TextMatrix(grid_unid.Row, WSCOL + 1) = Format(valor, "0.000")
  
    If cbo_ViewMoney.ListIndex = 1 Then ''Tag = "D" Then
      grid_unid.TextMatrix(grid_unid.Row, 11 + WSCOL) = Format(valor, "0.000")
    Else
      grid_unid.TextMatrix(grid_unid.Row, 16 + WSCOL) = Format(valor, "0.000")
    End If
End Sub
Public Sub CALCULAR_POR1(WSPRE As Currency, WSCOL As Long, ByVal iRow As Long)
Dim sPLH As Double
Dim sPorH As Double

    sPLH = WSPRE 'grid_unid.TextMatrix(grid_unid.Row, WSCOL)
    sPLH = (sPLH) / (1 + LK_IGV / 100)
    sPLH = Format(sPLH, "#0.000")
    sPorH = grid_unid.TextMatrix(iRow, 19)
    sPorH = ((sPorH - sPLH) / sPorH) * 100
     grid_unid.TextMatrix(iRow, WSCOL - 1) = Format(sPorH, "0.000")
    
End Sub
Public Sub CALCULAR_PRE1(WSPOR As Currency, WSCOL As Long)
Dim sPLH As Double
    
    sPLH = grid_unid.TextMatrix(CDbl(grid_unid.Row), 19)
    sPLH = sPLH - (sPLH * WSPOR / 100)
    sPLH = sPLH + (sPLH * LK_IGV / 100)
    grid_unid.TextMatrix(grid_unid.Row, WSCOL + 1) = Format(sPLH, "0.000")
        
End Sub
Private Sub cmdgrabar_Click()
Dim NroRows As Integer
Dim iRow As Long
On Error GoTo ErrorGrave
NroRows = grid_unid.Rows
If FLAG = 1 Then
    PROGRESO.Visible = True
    PROGRESO.max = NroRows
    For iRow = 1 To NroRows - 1
        PROGRESO.Value = iRow
        If Trim(grid_unid.TextMatrix(iRow, 44)) = "" Then
           PSPRE_LLAVE(0) = LK_CODCIA
           PSPRE_LLAVE(1) = grid_unid.TextMatrix(iRow, 46)
           PSPRE_LLAVE(2) = 0 'grid_unid.TextMatrix(iRow, 47)
           pre_llave.Requery
           pre_llave.Edit
           
           pre_llave!pre_pre11 = dDouble(grid_unid.TextMatrix(iRow, ColPre1 + 11))
           pre_llave!PRE_PRE22 = dDouble(grid_unid.TextMatrix(iRow, ColPre2 + 11))
           pre_llave!PRE_PRE33 = dDouble(grid_unid.TextMatrix(iRow, ColPre3 + 11))
           pre_llave!PRE_PRE44 = dDouble(grid_unid.TextMatrix(iRow, ColPre4 + 11))
           pre_llave!PRE_PRE55 = dDouble(grid_unid.TextMatrix(iRow, ColPre5 + 11))
           pre_llave!PRE_PRE66 = dDouble(grid_unid.TextMatrix(iRow, ColPre6 + 11))
           
           pre_llave!PRE_PRE1 = dDouble(grid_unid.TextMatrix(iRow, ColPre1 + 16))
           pre_llave!PRE_PRE2 = dDouble(grid_unid.TextMatrix(iRow, ColPre2 + 16))
           pre_llave!PRE_PRE3 = dDouble(grid_unid.TextMatrix(iRow, ColPre3 + 16))
           pre_llave!PRE_PRE4 = dDouble(grid_unid.TextMatrix(iRow, ColPre4 + 16))
           pre_llave!PRE_PRE5 = dDouble(grid_unid.TextMatrix(iRow, ColPre5 + 16))
           pre_llave!PRE_PRE6 = dDouble(grid_unid.TextMatrix(iRow, ColPre6 + 16))
           
           pre_llave!PRE_PORDES1 = dDouble(grid_unid.TextMatrix(iRow, 8))
           pre_llave!PRE_PORDES2 = dDouble(grid_unid.TextMatrix(iRow, 9))
           pre_llave!PRE_PORDES3 = dDouble(grid_unid.TextMatrix(iRow, 10))
           pre_llave!PRE_PORDES4 = dDouble(grid_unid.TextMatrix(iRow, 11))
           
           pre_llave!PRE_LITRO = dDouble(grid_unid.TextMatrix(iRow, 48)) 'precio de lista proveedor
           pre_llave!PRE_COSTO_REPO = dDouble(grid_unid.TextMatrix(iRow, 17)) 'margen
           
           pre_llave.Update
'           SQ_OPER = 1
'           PUB_KEY = grid_unid.TextMatrix(iRow, 31)
'           pu_codcia = LK_CODCIA
'           LEER_ART_LLAVE
'           If Not art_LLAVE.EOF Then
'            art_LLAVE.Edit
'            art_LLAVE("Art_Alterno") = grid_unid.TextMatrix(iRow, 1)
'            art_LLAVE("ART_CUENTA_CONTAB_C") = grid_unid.TextMatrix(iRow, 3)
'            art_LLAVE.Update
'           End If
        ElseIf Trim(grid_unid.TextMatrix(iRow, 29)) = "D" Then
            PUB_CODART = grid_unid.TextMatrix(iRow, 31)
            'cmdEliminar_Click
        End If
    Next iRow
    'grid_unid.Clear
    'grid_unid.Visible = False
    PROGRESO.Value = 1
    PROGRESO.Visible = False
    MsgBox "Se Actualizaron los precios", vbInformation, Pub_Titulo
    cmdCalcular.Enabled = False
    cbo_ViewMoney.Enabled = False
Else
    MsgBox "No existen articulos para actualizar", vbInformation, Pub_Titulo
End If
Exit Sub
ErrorGrave:
    PROGRESO.Value = 1
    PROGRESO.Visible = False
    MsgBox Err.Description, vbCritical, Pub_Titulo
End Sub

Private Sub grid_unid_ChangeTextBox()
Dim iRow As Long
    If DescPre = 0 Then
        If grid_unid.COL = ColPre1 - 1 Or grid_unid.COL = ColPre2 - 1 Or grid_unid.COL = ColPre3 - 1 Or grid_unid.COL = ColPre4 - 1 Or grid_unid.COL = ColPre5 - 1 Or grid_unid.COL = ColPre6 - 1 Then ' costo PORCENTAJE
            CALCULAR_PRE Val(grid_unid.TEXTO), grid_unid.COL
        End If
        If grid_unid.COL = ColPre1 Or grid_unid.COL = ColPre2 Or grid_unid.COL = ColPre3 Or grid_unid.COL = ColPre4 Or grid_unid.COL = ColPre5 Or grid_unid.COL = ColPre6 Then ' costo PORCENTAJE
            CALCULAR_POR Val(grid_unid.TEXTO), grid_unid.COL
        End If
    ElseIf DescPre = 1 Then
        iRow = grid_unid.Row
        If grid_unid.COL = 14 Then 'PLP
            CalculaLinea iRow, dDouble(grid_unid.TEXTO), dDouble(grid_unid.TextMatrix(iRow, 17))
        ElseIf grid_unid.COL = 17 Then 'Margen
            CalculaLinea iRow, dDouble(grid_unid.TextMatrix(iRow, 6)), dDouble(grid_unid.TEXTO)
        End If
        If grid_unid.COL = ColPre1 - 1 Or grid_unid.COL = ColPre2 - 1 Or grid_unid.COL = ColPre3 - 1 Or grid_unid.COL = ColPre4 - 1 Or grid_unid.COL = ColPre5 - 1 Or grid_unid.COL = ColPre6 - 1 Then ' costo PORCENTAJE
            CALCULAR_PRE1 Val(grid_unid.TEXTO), grid_unid.COL
        End If
        If grid_unid.COL = ColPre1 Or grid_unid.COL = ColPre2 Or grid_unid.COL = ColPre3 Or grid_unid.COL = ColPre4 Or grid_unid.COL = ColPre5 Or grid_unid.COL = ColPre6 Then ' costo PORCENTAJE
            CALCULAR_POR1 Val(grid_unid.TEXTO), grid_unid.COL, grid_unid.Row
        End If
        
    End If
End Sub

Private Sub grid_UNID_DblClick()
Dim iCol As Long

  If grid_unid.COL = 1 Then
    If Trim(grid_unid.TextMatrix(grid_unid.Row, 44)) = "" Then
        grid_unid.TextMatrix(grid_unid.Row, 44) = "X"
        For iCol = 1 To 10
            grid_unid.COL = iCol
            grid_unid.CellForeColor = QBColor(8)
        Next iCol
    ElseIf Trim(grid_unid.TextMatrix(grid_unid.Row, 44)) = "X" Then
        grid_unid.TextMatrix(grid_unid.Row, 44) = ""
        For iCol = 1 To 10
            grid_unid.COL = iCol
            grid_unid.CellForeColor = QBColor(0)
        Next iCol
    End If
  End If
End Sub

Private Sub grid_unid_ExitCell(introw As Integer, intColumn As Integer)
    If intColumn >= 8 And intColumn <= 11 Then
        CalculaLinea introw, dDouble(grid_unid.TextMatrix(CLng(introw), 6)), dDouble(grid_unid.TextMatrix(CLng(introw), 17))
    End If
End Sub

Private Sub txtPorAume1_KeyDown(KeyCode As Integer, Shift As Integer)
    CalFactorAume
End Sub
Private Sub txtPorAume2_KeyDown(KeyCode As Integer, Shift As Integer)
    CalFactorAume
End Sub
Private Sub txtPorDesc1_KeyUp(KeyCode As Integer, Shift As Integer)
    CalFactorDesc
End Sub

Private Sub txtPorDesc1_LostFocus()
    Call LLenaColVal(txtPorDesc1.Text, 8)
    chkDesc2.SetFocus
End Sub

Private Sub txtPorDesc2_KeyUp(KeyCode As Integer, Shift As Integer)
    CalFactorDesc
End Sub

Private Sub txtPorDesc2_LostFocus()
    Call LLenaColVal(txtPorDesc2.Text, 9)
    If chkDesc3.Enabled Then
        chkDesc3.SetFocus
    End If
End Sub
Private Sub txtPorDesc3_KeyUp(KeyCode As Integer, Shift As Integer)
    CalFactorDesc
End Sub
Private Sub txtPorDesc3_LostFocus()
    Call LLenaColVal(txtPorDesc3.Text, 10)
    If chkDesc4.Enabled Then
        chkDesc4.SetFocus
    End If
End Sub
Private Sub txtPorDesc4_KeyUp(KeyCode As Integer, Shift As Integer)
    CalFactorDesc
End Sub
Private Sub txtPorDesc4_LostFocus()
    Call LLenaColVal(txtPorDesc4.Text, 11)
End Sub

Private Sub LLenaColVal(ByVal valor As Variant, ByVal iCol As Long)
Dim iRows As Integer
Dim iRow As Long
    iRows = grid_unid.Rows - 1
    If IsNumeric(valor) Then valor = Format(valor, "#0.000")
    For iRow = 1 To iRows
        grid_unid.TextMatrix(iRow, iCol) = valor
    Next iRow
End Sub
Private Sub ChangeeViewMoney()
Dim WSPOR As Currency
Dim money As String

    If cbo_ViewMoney.ListIndex = 0 Then ''.Tag = "D" Then
        money = "S" ''cbo_ViewMoney.Tag = "S"
    ElseIf cbo_ViewMoney.ListIndex = 1 Then
        money = "D"  ''cbo_ViewMoney.Tag = "D"
    End If
    
    For fila = 1 To grid_unid.Rows - 1
        If money = "D" Then ''cbo_ViewMoney.Tag = "D" Then
            grid_unid.TextMatrix(fila, ColPre1) = grid_unid.TextMatrix(fila, ColPre1 + 11)
            grid_unid.TextMatrix(fila, ColPre2) = grid_unid.TextMatrix(fila, ColPre2 + 11)
            grid_unid.TextMatrix(fila, ColPre3) = grid_unid.TextMatrix(fila, ColPre3 + 11)
            grid_unid.TextMatrix(fila, ColPre4) = grid_unid.TextMatrix(fila, ColPre4 + 11)
            grid_unid.TextMatrix(fila, ColPre5) = grid_unid.TextMatrix(fila, ColPre5 + 11)
            grid_unid.TextMatrix(fila, ColPre6) = grid_unid.TextMatrix(fila, ColPre6 + 11)
        Else
            grid_unid.TextMatrix(fila, ColPre1) = grid_unid.TextMatrix(fila, ColPre1 + 16)
            grid_unid.TextMatrix(fila, ColPre2) = grid_unid.TextMatrix(fila, ColPre2 + 16)
            grid_unid.TextMatrix(fila, ColPre3) = grid_unid.TextMatrix(fila, ColPre3 + 16)
            grid_unid.TextMatrix(fila, ColPre4) = grid_unid.TextMatrix(fila, ColPre4 + 16)
            grid_unid.TextMatrix(fila, ColPre5) = grid_unid.TextMatrix(fila, ColPre5 + 16)
            grid_unid.TextMatrix(fila, ColPre6) = grid_unid.TextMatrix(fila, ColPre6 + 16)
        End If
        If money = "D" Then ''cbo_ViewMoney.Tag = "D" Then
            grid_unid.TextMatrix(fila, 6) = redondea(Val(grid_unid.TextMatrix(fila, 48)) / LK_TIPO_CAMBIO)
        Else
            grid_unid.TextMatrix(fila, 6) = redondea(Val(grid_unid.TextMatrix(fila, 48)))
        End If
     Next fila
     If FlagPrecio = 2 Then cmdCalcular_Click
End Sub
Private Sub BackColorRow(ByVal iRow As Long)
Dim iCol As Long
    grid_unid.Row = iRow
    For iCol = 1 To 31
     grid_unid.COL = iCol
     grid_unid.CellBackColor = &H80000018
     'grid_unid.CellFontBold = True
    Next
End Sub

Private Sub SETGRID()
    grid_unid.FormatString = "Descripcion|Codigo|Unidad|Equiv.|||PreLisProv||Por1|Por2|Por3|Por4||FactDesc||CVV|CVV/IGV|Margen|Res|PLH|(%)|Prec1|(%)|Prec2|(%)|Prec3|(%)|Prec14|(%)|Prec5|(%)|Prec6"
    grid_unid.Rows = 2
    grid_unid.Cols = 49

    grid_unid.ColWidth(0) = 3800 ' Descripcion
    grid_unid.ColWidth(1) = 800 'Codigo
    grid_unid.ColWidth(2) = 800 'Unidad
    grid_unid.ColWidth(3) = 800 ' equiv
    grid_unid.ColWidth(4) = 0  '
    grid_unid.ColWidth(5) = 0  '
    grid_unid.ColWidth(6) = 800  'PrecLista Proveedor
    grid_unid.ColWidth(7) = 0  '
    grid_unid.ColWidth(8) = 800  'PorDesc1
    grid_unid.ColWidth(9) = 0  'PorDesc2
    grid_unid.ColWidth(10) = 0  'PorDesc3
    grid_unid.ColWidth(11) = 0  'PorDesc4
    grid_unid.ColWidth(12) = 0  '
    grid_unid.ColWidth(13) = 800  'Factor de Descuento
    
    grid_unid.ColWidth(14) = 0  '
    grid_unid.ColWidth(15) = 800  'CVV
    grid_unid.ColWidth(16) = 800  'CVV/IGV
    
    grid_unid.ColWidth(17) = 800  'Margen
    grid_unid.ColWidth(18) = 700  'Res
    grid_unid.ColWidth(19) = 840  ' PLH
    grid_unid.ColWidth(20) = 700   ' % 1
    grid_unid.ColWidth(21) = 950   ' p 1
    grid_unid.ColWidth(22) = 700   ' % 2
    grid_unid.ColWidth(23) = 950   ' p 2
    grid_unid.ColWidth(24) = 700   ' % 3
    grid_unid.ColWidth(25) = 950   ' p 3
    grid_unid.ColWidth(26) = 700   ' % 4
    grid_unid.ColWidth(27) = 950   ' p 4
    grid_unid.ColWidth(28) = 700   ' % 5
    grid_unid.ColWidth(29) = 950   ' p 5
    grid_unid.ColWidth(30) = 700   ' % 6
    grid_unid.ColWidth(31) = 950   ' p 6
    grid_unid.ColWidth(32) = 1   ' Guarda p.digitado 11(D)
    grid_unid.ColWidth(33) = 1   ' Guarda p.digitado 22(D)
    grid_unid.ColWidth(34) = 1   ' Guarda p.digitado 33(D)
    grid_unid.ColWidth(35) = 1   ' Guarda p.digitado 44(D)
    grid_unid.ColWidth(36) = 1   ' Guarda p.digitado 55(D)
    grid_unid.ColWidth(37) = 1   ' Guarda p.digitado 66(D)
    grid_unid.ColWidth(38) = 1   ' Guarda p.digitado 1(S)
    grid_unid.ColWidth(39) = 1   ' Guarda p.digitado 2(S)
    grid_unid.ColWidth(40) = 1   ' Guarda p.digitado 3(S)
    grid_unid.ColWidth(41) = 1   ' Guarda p.digitado 4(S)
    grid_unid.ColWidth(42) = 1   ' Guarda p.digitado 5(S)
    grid_unid.ColWidth(43) = 1   ' Guarda p.digitado 6(S)
    grid_unid.ColWidth(44) = 1   ' Guarda si graba o no
    grid_unid.ColWidth(45) = 1   ' Guarda el costo Rep. soles
    grid_unid.ColWidth(46) = 1   ' Guarda keyarticulo
    grid_unid.ColWidth(47) = 1   ' Guarda secuencia
    grid_unid.ColWidth(48) = 1   ' Guarda Prec Lista Proveed en Soles
   
    grid_unid.ColControl(6) = osTextBox 'PLP
    grid_unid.ColEdit(6) = True
    grid_unid.ColData(6) = osDecimal
    grid_unid.ColSalto(6) = 2
    
    grid_unid.ColControl(8) = osTextBox 'Porc de Descto 1
    grid_unid.ColEdit(8) = True
    grid_unid.ColData(8) = osDecimal
    grid_unid.ColSalto(8) = 9
    
    grid_unid.ColControl(17) = osTextBox 'Codigo
    grid_unid.ColEdit(17) = True
    grid_unid.ColData(17) = osDecimal
    grid_unid.ColSalto(17) = 4
    
    grid_unid.ColControl(ColPre1 - 1) = osTextBox '  % Pre1
    grid_unid.ColEdit(ColPre1 - 1) = True
    grid_unid.ColData(ColPre1 - 1) = osDecimal
    grid_unid.ColSalto(ColPre1 - 1) = 1
    
    grid_unid.ColControl(ColPre1) = osTextBox ' Pre1
    grid_unid.ColEdit(ColPre1) = True
    grid_unid.ColData(ColPre1) = osDecimal
    grid_unid.ColSalto(ColPre1) = 1
    
    grid_unid.ColControl(ColPre2 - 1) = osTextBox ' % Pre2
    grid_unid.ColEdit(ColPre2 - 1) = True
    grid_unid.ColData(ColPre2 - 1) = osDecimal
    grid_unid.ColSalto(ColPre2 - 1) = 1
    
    grid_unid.ColControl(ColPre2) = osTextBox ' Pre2
    grid_unid.ColEdit(ColPre2) = True
    grid_unid.ColData(ColPre2) = osDecimal
    grid_unid.ColSalto(ColPre2) = 1
    
    grid_unid.ColControl(ColPre3 - 1) = osTextBox ' % Pre3
    grid_unid.ColEdit(ColPre3 - 1) = True
    grid_unid.ColData(ColPre3 - 1) = osDecimal
    grid_unid.ColSalto(ColPre3 - 1) = 1
    
    grid_unid.ColControl(ColPre3) = osTextBox ' Pre3
    grid_unid.ColEdit(ColPre3) = True
    grid_unid.ColData(ColPre3) = osDecimal
    grid_unid.ColSalto(ColPre3) = 1
    
    grid_unid.ColControl(ColPre4 - 1) = osTextBox ' % Pre4
    grid_unid.ColEdit(ColPre4 - 1) = True
    grid_unid.ColData(ColPre4 - 1) = osDecimal
    grid_unid.ColSalto(ColPre4 - 1) = 1
    
    grid_unid.ColControl(ColPre4) = osTextBox ' % Pre4
    grid_unid.ColEdit(ColPre4) = True
    grid_unid.ColData(ColPre4) = osDecimal
    grid_unid.ColSalto(ColPre4) = 1
    
    grid_unid.ColControl(ColPre5 - 1) = osTextBox ' % Pre5
    grid_unid.ColEdit(ColPre5 - 1) = True
    grid_unid.ColData(ColPre5 - 1) = osDecimal
    grid_unid.ColSalto(ColPre5 - 1) = 1
    
    grid_unid.ColControl(ColPre5) = osTextBox ' Pre5
    grid_unid.ColEdit(ColPre5) = True
    grid_unid.ColData(ColPre5) = osDecimal
    grid_unid.ColSalto(ColPre5) = 1

    grid_unid.ColControl(ColPre6 - 1) = osTextBox ' % Pre6
    grid_unid.ColEdit(ColPre6 - 1) = True
    grid_unid.ColData(ColPre6 - 1) = osDecimal
    grid_unid.ColSalto(ColPre6 - 1) = 1
    
    grid_unid.ColControl(ColPre6) = osTextBox ' Pre6
    grid_unid.ColEdit(ColPre6) = True
    grid_unid.ColData(ColPre6) = osDecimal
    grid_unid.ColSalto(ColPre6) = 1
    
    BackColorRow 1
End Sub

Private Sub cmdParametros_Click()
    frmParametros.Visible = Not frmParametros.Visible
'    CalFactorDesc
    dFactorManejo = dDouble(txtFactorManejo.Text)

End Sub

Private Sub Form_Load()
    LlenadoCbo art_familia, 122
    LlenadoCbo art_grupo, 129
    LlenadoCbo art_numero, 130
    SETGRID
    cbo_Unidad.ListIndex = 0
    txtFactorManejo.Text = LK_FACTORMAN
    dFactorManejo = LK_FACTORMAN
    DescSuc = 1
    DescPre = 1
End Sub
Private Sub cbo_ViewMoney_Click()
If Not grid_unid.Rows > 1 Then Exit Sub
If cbo_ViewMoney.ListIndex = -1 Then Exit Sub
    ChangeeViewMoney
End Sub

Private Sub chkDesc_Click()
    If chkDesc.Value = 1 Then
        DescSuc = 1
    Else
        DescSuc = 0
        FactorDesc = 0
        lblFactor.Caption = 0
    End If
End Sub

Private Sub chkDescPre_Click()
    If chkDesc.Value = 1 Then
        DescPre = 1
    Else
        DescPre = 0
    End If
End Sub
