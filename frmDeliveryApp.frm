VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form frmDeliveryApp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delivery"
   ClientHeight    =   9870
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   16050
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9870
   ScaleWidth      =   16050
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvCliente 
      Height          =   2505
      Left            =   2280
      TabIndex        =   48
      Top             =   1005
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4419
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   1440
      TabIndex        =   67
      Top             =   2400
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdRepartidor 
      Caption         =   "Repartidor"
      Height          =   360
      Left            =   14400
      TabIndex        =   61
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdPagar 
      Caption         =   "&Pagar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   14880
      TabIndex        =   55
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdEnviarDelivery 
      Caption         =   "Facturar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   13680
      TabIndex        =   54
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdClienteEdit 
      Caption         =   "Cambiar Cliente"
      Height          =   360
      Left            =   5400
      TabIndex        =   53
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdFormaPago 
      Caption         =   "Forma de Pago"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   12480
      TabIndex        =   49
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdConsulta 
      Caption         =   "Consulta Últimas Compras"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   11160
      TabIndex        =   47
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame FraObs 
      Height          =   2415
      Left            =   7320
      TabIndex        =   43
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txtObs 
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   44
         Top             =   480
         Width           =   3495
      End
      Begin VB.CommandButton cmdActualizaObservaciones 
         Caption         =   "Actualiza Observaciones"
         Height          =   360
         Left            =   1680
         TabIndex        =   57
         Top             =   120
         Width           =   2310
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBS"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame FraPedido 
      Height          =   7215
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   15855
      Begin VB.CommandButton cmdDescuentos 
         Caption         =   "Descuentos"
         Height          =   480
         Left            =   5760
         TabIndex        =   68
         Top             =   3360
         Width           =   1110
      End
      Begin VB.CommandButton cmdCaracteristicas 
         Caption         =   "Caracteristicas"
         Height          =   495
         Left            =   10800
         TabIndex        =   60
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton cmdPorcion 
         Caption         =   "1/2 Porcion"
         Height          =   495
         Left            =   10800
         TabIndex        =   59
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdDoctoDespacho 
         BackColor       =   &H0000FFFF&
         Caption         =   "Dcto. Despacho"
         Height          =   495
         Left            =   10800
         TabIndex        =   56
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdEnviar 
         Caption         =   "&Imprimir"
         Height          =   495
         Left            =   10800
         Picture         =   "frmDeliveryApp.frx":0000
         TabIndex        =   42
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "&Detalle"
         Height          =   495
         Left            =   10800
         Picture         =   "frmDeliveryApp.frx":08CA
         TabIndex        =   41
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   495
         Left            =   10800
         Picture         =   "frmDeliveryApp.frx":1194
         TabIndex        =   40
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdBorrar 
         Height          =   735
         Left            =   14760
         TabIndex        =   36
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdLimpiar 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   14760
         TabIndex        =   37
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdPrecio 
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   14760
         TabIndex        =   38
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdCantidad 
         Caption         =   "Cant"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   14760
         TabIndex        =   39
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdPunto 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   13800
         TabIndex        =   35
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   11880
         TabIndex        =   25
         Top             =   3120
         Width           =   1935
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   13800
         TabIndex        =   28
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   12840
         TabIndex        =   27
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   11880
         TabIndex        =   26
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   6
         Left            =   13800
         TabIndex        =   31
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   12840
         TabIndex        =   30
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   11880
         TabIndex        =   29
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   9
         Left            =   13800
         TabIndex        =   34
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   8
         Left            =   12840
         TabIndex        =   33
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   11880
         TabIndex        =   32
         Top             =   960
         Width           =   975
      End
      Begin VB.Frame fraPlatos 
         Height          =   3105
         Left            =   9100
         TabIndex        =   13
         Top             =   3840
         Width           =   6630
         Begin VB.CommandButton cmdPlatoSig 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   5520
            Picture         =   "frmDeliveryApp.frx":1A5E
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   2325
            Width           =   1095
         End
         Begin VB.CommandButton cmdPlato 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   0
            Left            =   5400
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   720
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdPlatoAnt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   30
            Picture         =   "frmDeliveryApp.frx":2328
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame fraSubFamilia 
         Height          =   3105
         Left            =   4150
         TabIndex        =   12
         Top             =   3840
         Width           =   4900
         Begin VB.CommandButton cmdSubFam 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   0
            Left            =   960
            TabIndex        =   19
            Top             =   1320
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton cmdSubFamSig 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3675
            Picture         =   "frmDeliveryApp.frx":2BF2
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   2325
            Width           =   1215
         End
         Begin VB.CommandButton cmdSubFamAnt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   30
            Picture         =   "frmDeliveryApp.frx":34BC
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Frame FraFamilia 
         Height          =   3105
         Left            =   120
         TabIndex        =   11
         Top             =   3840
         Width           =   3975
         Begin VB.CommandButton cmdFamSig 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   2940
            Picture         =   "frmDeliveryApp.frx":3D86
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   2325
            Width           =   975
         End
         Begin VB.CommandButton cmdFam 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   0
            Left            =   1440
            TabIndex        =   15
            Top             =   1800
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cmdFamAnt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   30
            Picture         =   "frmDeliveryApp.frx":4650
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   120
            Width           =   975
         End
      End
      Begin MSComctlLib.ListView lvPlatos 
         Height          =   3135
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ilPedido"
         SmallIcons      =   "ilPedido"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblicbperORI 
         Caption         =   "0"
         Height          =   255
         Left            =   2880
         TabIndex        =   74
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label lblDni 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "dni"
         Height          =   195
         Left            =   7200
         TabIndex        =   73
         Top             =   3480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblicbper 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   4200
         TabIndex        =   69
         Top             =   3480
         Width           =   105
      End
      Begin VB.Label lblItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
         Height          =   195
         Left            =   240
         TabIndex        =   52
         Top             =   3480
         Width           =   555
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   435
         Left            =   11880
         TabIndex        =   46
         Top             =   360
         Width           =   3795
      End
      Begin VB.Label lblTot 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   8640
         TabIndex        =   24
         Top             =   3480
         Width           =   2070
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   7920
         TabIndex        =   23
         Top             =   3480
         Width           =   660
      End
   End
   Begin VB.Frame FraCliente 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.CheckBox chkRecojo 
         Caption         =   "Delivery"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2880
         TabIndex        =   70
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtUrb 
         Height          =   290
         Left            =   1320
         TabIndex        =   66
         Top             =   1920
         Width           =   2655
      End
      Begin MSDataListLib.DataCombo DatDireccion 
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   960
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.CommandButton cmdCliente 
         Caption         =   "..."
         Height          =   290
         Left            =   6240
         TabIndex        =   7
         Top             =   960
         Width           =   615
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "TELEFONO"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "CLIENTE"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   3
         Top             =   270
         Value           =   -1  'True
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo datZona 
         Height          =   315
         Left            =   4680
         TabIndex        =   64
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtCliente 
         Height          =   290
         Left            =   2160
         TabIndex        =   2
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label lblReferencia 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   525
         Left            =   1320
         TabIndex        =   72
         Top             =   1320
         Width           =   4890
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REFERENCIA:"
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   1320
         Width           =   1170
      End
      Begin VB.Label lblUrb 
         BackStyle       =   0  'Transparent
         Caption         =   "-1"
         Height          =   195
         Left            =   6240
         TabIndex        =   65
         Top             =   1920
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ZONA:"
         Height          =   195
         Left            =   4080
         TabIndex        =   63
         Top             =   1980
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "URB:"
         Height          =   195
         Left            =   720
         TabIndex        =   62
         Top             =   1920
         Width           =   435
      End
      Begin VB.Label lblruc 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   3360
         TabIndex        =   58
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIRECCIÓN:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1020
         Width           =   1110
      End
      Begin VB.Label lblCliente 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLIENTE:"
         Height          =   195
         Left            =   420
         TabIndex        =   1
         Top             =   645
         Width           =   810
      End
   End
   Begin MSComctlLib.ImageList ilPedido 
      Left            =   14760
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeliveryApp.frx":4F1A
            Key             =   "Plato"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblnumero 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   11640
      TabIndex        =   51
      Top             =   840
      Width           =   2715
   End
   Begin VB.Label lblserie 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   11640
      TabIndex        =   50
      Top             =   240
      Width           =   2715
   End
End
Attribute VB_Name = "frmDeliveryApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private loc_key_u  As Integer
Private vBuscar_u As Boolean 'variable para la busqueda de clientes
Private vPagActFam, vPagActSubFam, vPagActPla As Integer
Private vPagTotFam, vPagTotSubFam, vPagTotPla As Integer
Private oRsFam As ADODB.Recordset
Private oRsSubFam As ADODB.Recordset
Public oRsPlatos As ADODB.Recordset
Private vIniLeft As Integer
Private vIniTop As Integer
Public vCodFam As Integer
Private vFiltro As Integer 'variable para saber porque opcion se esta bsucando
Private loc_key As Integer
Private vBuscar As Boolean 'variable para la busqueda de clientes
Public VNuevo As Boolean
Public vnumser As String
Public vNumFac As Double
Public vPrimero As Boolean
Public vEstado As String
Public vMaxFac As Double
Private cMontoTarifa As Double
Private cMontoDescuento As Double
Private ORSurb As ADODB.Recordset
Private oRSdir As ADODB.Recordset
Public gEstaCargando As Boolean         '16-06-2020

Private Sub CambiaPrecio()

    If Not IsNumeric(Me.lblTexto.Caption) Then
        MsgBox "No ha ingresado ningún precio"

        Exit Sub

    End If

    If Not Me.lvPlatos.SelectedItem Is Nothing Then
        LimpiaParametros oCmdEjec

        With oCmdEjec
            .CommandText = "SpModificarPreCantPla"
            .Parameters.Append .CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
            .Parameters.Append .CreateParameter("@NumSer", adChar, adParamInput, 4, Trim(Me.lblSerie.Caption))
            .Parameters.Append .CreateParameter("@NumFac", adDouble, adParamInput, , CDbl(Me.lblNumero.Caption))
            .Parameters.Append .CreateParameter("@NumSec", adInteger, adParamInput, , CInt(Me.lvPlatos.SelectedItem.SubItems(6)))
            .Parameters.Append .CreateParameter("@CodArt", adDouble, adParamInput, , CDbl(Me.lvPlatos.SelectedItem.Tag))
            .Parameters.Append .CreateParameter("@Pre", adDouble, adParamInput, , CDbl(Me.lblTexto.Caption))
            .Parameters.Append .CreateParameter("@Cant", adInteger, adParamInput, , Null)

            .Parameters.Append .CreateParameter("@EsPre", adBoolean, adParamInput, , True)
            .Execute
        End With

        Me.lvPlatos.SelectedItem.SubItems(4) = FormatNumber(Me.lblTexto.Caption, 2)
        Me.lvPlatos.SelectedItem.SubItems(5) = FormatNumber(val(Me.lvPlatos.SelectedItem.SubItems(3)) * val(Me.lvPlatos.SelectedItem.SubItems(4)), 2)
        Me.lblTot.Caption = FormatCurrency(sumatoria, 2)
        Me.lblTexto.Caption = ""
    End If
 
    oCmdEjec.CommandText = "USP_COMANDA_ACTUALIZA_PAGOSDELIVERY"
    LimpiaParametros oCmdEjec
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adVarChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adDouble, adParamInput, , Me.lblNumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TOTAL", adDouble, adParamInput, , Me.lblTot.Caption)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adDouble, adParamInput, , Me.DatZona.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adDouble, adParamInput, , Me.DatZona.BoundText)
    
    oCmdEjec.Execute
 
End Sub

Private Function VerificaPass(vUSUARIO As String, vClave As String, ByRef vMSN As String) As Boolean
Dim orsPass As ADODB.Recordset
Dim vtpass As String, vPasa As Boolean
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SpDevuelveClaveCaja"
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, vUSUARIO)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CLAVE", adVarChar, adParamInput, 10, vClave)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MSN", adVarChar, adParamOutput, 200, 1)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)
oCmdEjec.Execute

'If Not orsPass.EOF Then vtpass = Trim(orsPass!Clave)
vtpass = oCmdEjec.Parameters("@MSN").Value
vPasa = oCmdEjec.Parameters("@PASA").Value
vMSN = vtpass

    VerificaPass = vPasa
End Function

Private Sub elimina()

    On Error GoTo eli

    Pub_ConnAdo.BeginTrans
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpActualizarPlato"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adVarChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adDouble, adParamInput, , Me.lblNumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numsec", adInteger, adParamInput, , Me.lvPlatos.SelectedItem.SubItems(6))
    oCmdEjec.Execute
    
    'actualiza stock
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpActualizaStock"

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Usuario", adVarChar, adParamInput, 20, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodArt", adBigInt, adParamInput, , Me.lvPlatos.SelectedItem.Tag)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@cp", adBigInt, adParamInput, , Me.lvPlatos.SelectedItem.SubItems(3))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ser", adChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@nro", adInteger, adParamInput, , Me.lblNumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@mesa", adChar, adParamInput, 10, "")
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@tipo", adBoolean, adParamInput, , 0)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumSec", adInteger, adParamInput, , 0)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MaxNumFac", adInteger, adParamOutput, , 3)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MaxNumOper", adInteger, adParamOutput, , 3)

    oCmdEjec.Execute
    
   
    
    Dim XDTA As Double

    XDTA = CDbl(Me.lblNumero.Caption)    'linea nueva
    If Me.lvPlatos.ListItems.count = 1 Then
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SpLiberarDelivery"
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
        
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adInteger, adParamInput, , Me.lblSerie.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , XDTA)
        oCmdEjec.Execute

  oCmdEjec.CommandText = "USP_COMANDA_ACTUALIZA_PAGOSDELIVERY"
    LimpiaParametros oCmdEjec
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adVarChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adDouble, adParamInput, , Me.lblNumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TOTAL", adDouble, adParamInput, , Me.lblTot.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adDouble, adParamInput, , Me.DatZona.BoundText)
    
    oCmdEjec.Execute
        Unload Me
    
    Else
        Me.lvPlatos.ListItems.Remove Me.lvPlatos.SelectedItem.index
        Me.lblTot.Caption = FormatCurrency(sumatoria, 2)
        Me.lblItems.Caption = "Items: " & Me.lvPlatos.ListItems.count
        
          oCmdEjec.CommandText = "USP_COMANDA_ACTUALIZA_PAGOSDELIVERY"
    LimpiaParametros oCmdEjec
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adVarChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adDouble, adParamInput, , Me.lblNumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TOTAL", adDouble, adParamInput, , Me.lblTot.Caption)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adDouble, adParamInput, , Me.DatZona.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adDouble, adParamInput, , Me.DatZona.BoundText)
    
    oCmdEjec.Execute
       
    End If


'ACTUALIZA LA COMANDA A FACTURADA Y LIBERA LA MESA CUANDO SE EXTORNA UN PLATO
    oCmdEjec.CommandText = "SP_PEDIDO_FACTURADO"
    LimpiaParametros oCmdEjec
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , XDTA)
        
    oCmdEjec.Execute
    
   
    
    Pub_ConnAdo.CommitTrans

    Exit Sub

eli:
    Pub_ConnAdo.RollbackTrans
    MsgBox Err.Description
End Sub

Private Sub extorna(cIDmotivo As Integer, cMOTIVO As String, cUSUARIO As String)

    On Error GoTo ext

    Pub_ConnAdo.BeginTrans
    
    'si el idMOTIVO ES -1 ENTONCES ES UN MOTIVO NUEVO Y LO REGISTRA EN EL MAESTRO DE MOTIVOS
    If cIDmotivo = -1 Then

        Dim tIDmotivo As Integer
        
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SP_MOTIVOS_REGISTRAR"
    
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DESCRIPCION", adVarChar, adParamInput, 100, cMOTIVO)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDMOTIVO", adInteger, adParamOutput, , tIDmotivo)
        oCmdEjec.Execute
    
        tIDmotivo = oCmdEjec.Parameters("@IDMOTIVO").Value
        cIDmotivo = tIDmotivo
    End If
    
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpActualizarPlato1"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adVarChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adDouble, adParamInput, , Me.lblNumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numsec", adInteger, adParamInput, , Me.lvPlatos.SelectedItem.SubItems(6))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Usuario", adVarChar, adParamInput, 20, frmClaveCaja.vUSUARIO)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDMOTIVO", adInteger, adParamInput, , cIDmotivo)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIOELIMINA", adVarChar, adParamInput, 10, cUSUARIO)
    oCmdEjec.Execute
    
    'actualiza stock
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpActualizaStock"

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Usuario", adVarChar, adParamInput, 20, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodArt", adBigInt, adParamInput, , Me.lvPlatos.SelectedItem.Tag)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@cp", adBigInt, adParamInput, , Me.lvPlatos.SelectedItem.SubItems(3))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ser", adChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@nro", adInteger, adParamInput, , Me.lblNumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@mesa", adChar, adParamInput, 10, vMesa)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@tipo", adBoolean, adParamInput, , 0)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumSec", adInteger, adParamInput, , 0)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MaxNumFac", adInteger, adParamOutput, , 3)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MaxNumOper", adInteger, adParamOutput, , 3)

    oCmdEjec.Execute
    Dim XDTA As Double
    XDTA = CDbl(Me.lblNumero.Caption) 'NUMERO DE COMANDA
    
 If Me.lvPlatos.ListItems.count = 1 Then
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SpLiberarDelivery"
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
        
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adInteger, adParamInput, , Me.lblSerie.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , XDTA)
        oCmdEjec.Execute
        
         oCmdEjec.CommandText = "USP_COMANDA_ACTUALIZA_PAGOSDELIVERY"
    LimpiaParametros oCmdEjec
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adVarChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adDouble, adParamInput, , Me.lblNumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TOTAL", adDouble, adParamInput, , Me.lblTot.Caption)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adDouble, adParamInput, , Me.DatZona.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adDouble, adParamInput, , Me.DatZona.BoundText)
    oCmdEjec.Execute

        Unload Me
    Else
        Me.lvPlatos.ListItems.Remove Me.lvPlatos.SelectedItem.index
        Me.lblTot.Caption = FormatCurrency(sumatoria, 2)
        Me.lblItems.Caption = "Items: " & Me.lvPlatos.ListItems.count
        
         oCmdEjec.CommandText = "USP_COMANDA_ACTUALIZA_PAGOSDELIVERY"
    LimpiaParametros oCmdEjec
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adVarChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adDouble, adParamInput, , Me.lblNumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TOTAL", adDouble, adParamInput, , Me.lblTot.Caption)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adDouble, adParamInput, , Me.DatZona.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adDouble, adParamInput, , Me.DatZona.BoundText)
    oCmdEjec.Execute
    End If
    
    
    'ACTUALIZA LA COMANDA A FACTURADA Y LIBERA LA MESA CUANDO SE EXTORNA UN PLATO
    oCmdEjec.CommandText = "SP_PEDIDO_FACTURADO"
    LimpiaParametros oCmdEjec
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , XDTA)
        
    oCmdEjec.Execute
    
    
    
    
    Pub_ConnAdo.CommitTrans

    Exit Sub

ext:
    Pub_ConnAdo.RollbackTrans
    MsgBox Err.Description
End Sub

Public Function sumatoria() As Currency
Dim fila As ListItem
Dim vTot As Currency
vTot = 0

For c = 1 To Me.lvPlatos.ListItems.count
    vTot = vTot + val(Me.lvPlatos.ListItems(c).SubItems(5))
Next

'For Each fila In Me.lvPlatos.ListItems
'    vTot = vTot + val(fila.SubItems(4))
'Next
sumatoria = vTot
End Function

Public Sub CargarComanda(vCodCia As String, vnumser As String, vNumFac As Double)

    Dim oRsComanda As ADODB.Recordset

    Me.lvPlatos.ListItems.Clear

    Dim vMozo    As String

    Dim vCodMozo As Integer

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPCARGARCOMANDADELIVERY"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adChar, adParamInput, 3, vnumser)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , vNumFac)

    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    Dim XfECHA As String

    XfECHA = CStr(Year(LK_FECHA_DIA))
    XfECHA = XfECHA & "-" & Right("00" & Month(LK_FECHA_DIA), 2)
    XfECHA = XfECHA & "-" & Right("00" & Day(LK_FECHA_DIA), 2)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Fecha", adDBTimeStamp, adParamInput, , XfECHA)
    
    Set oRsComanda = oCmdEjec.Execute

    If Not oRsComanda.EOF Then
        Me.lblNumero.Caption = oRsComanda.Fields!PED_NUMFAC 'oCmdEjec.Parameters("@NumFac").Value
        Me.lblSerie.Caption = oRsComanda.Fields!PED_numser 'oCmdEjec.Parameters("@NumSer").Value
        ' Me.lblMozo.Tag = oRsComanda.Fields!PED_CODVEN  'oCmdEjec.Parameters("@CodMozo").Value
        ' Me.lblMozo.Caption = Trim(oRsComanda.Fields!mozo) 'Trim(oCmdEjec.Parameters("@Mozo").Value)
        Me.lblCliente.Caption = IIf(IsNull(Trim(oRsComanda!cliente)), "", oRsComanda!cliente)

        ' Me.lblComensales.Caption = Trim(oRsComanda!Comensales)
        '    Me.lblSerie.Tag = oRsComanda!NumFac
        Dim cICBPEr As Double
        cICBPEr = 0
        Do While Not oRsComanda.EOF
    
            With Me.lvPlatos.ListItems.Add(, , Trim(oRsComanda!plato), Me.ilPedido.ListImages(1).key, Me.ilPedido.ListImages(1).key)
                .Tag = oRsComanda!CODPLATO
                '.SubItems(1) = iif(oRsComanda!cuenta Trim(oRsComanda!cuenta)
                .SubItems(1) = IIf(IsNull(oRsComanda!cuenta), "", Trim(oRsComanda!cuenta))
                .SubItems(2) = Trim(oRsComanda!DETALLE)
                .SubItems(3) = Format(oRsComanda!Cantidad, "#####0.#0")
                .SubItems(4) = Format(oRsComanda!PRECIO, "#####0.#0")
                .SubItems(5) = Format(oRsComanda!Importe, "#####0.#0")
                .SubItems(6) = oRsComanda!Sec
                .SubItems(7) = oRsComanda!aten
                '.SubItems(7) = oRsComanda!NumFac
                .SubItems(8) = oRsComanda.Fields!PED_NUMFAC
                .SubItems(9) = oRsComanda!aPRO
                .SubItems(10) = oRsComanda!NumFac
                If CBool(oRsComanda!icbper) Then
                'cICBPEr = cICBPEr + (oRsComanda!Cantidad * oRsComanda!gen_icbper)
                cICBPEr = cICBPEr + (oRsComanda!BOLSAS * oRsComanda!gen_icbper)
                Else
                cICBPEr = cICBPEr + oRsComanda!combo_icbper
                End If
                '.SubItems(10) = oRsComanda!PED_NUMFAC
                If oRsComanda!aPRO = "0" Then .Checked = True
            End With

            Me.lblTot.Caption = FormatCurrency(sumatoria, 2)
        
            oRsComanda.MoveNext
        Loop
        Me.lblicbper.Caption = cICBPEr
        Me.lblicbperORI.Caption = cICBPEr
        If Not VNuevo Then

            Dim orsP As ADODB.Recordset

            Set orsP = oRsComanda.NextRecordset
            Me.lblCliente.Caption = orsP!IDECLIENTE
            Me.txtCliente.Text = Trim(orsP!cliente)
            Me.lblRUC.Caption = orsP!RUC
            Me.lblDNI.Caption = orsP!DNI
            'OBS
            Me.txtObs.Text = orsP!OBS
            Me.chkRecojo.Value = orsP!RECOJO
            
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SP_DELIVERY_CLIENTE_DIRECCIONES"
            
            Me.DatDireccion.BoundText = ""

            Set oRSdir = oCmdEjec.Execute(, Array(LK_CODCIA, Me.lblCliente.Caption))
            Set Me.DatDireccion.RowSource = oRSdir
            Me.DatDireccion.BoundColumn = oRSdir.Fields(4).Name
            Me.DatDireccion.ListField = oRSdir.Fields(0).Name
        
            
            'AQUI
            
            Set ORSurb = oRSdir.NextRecordset
            Me.txtUrb.Text = ORSurb!urb
            Me.lblurb.Caption = ORSurb!IDEURB
            
            Dim ORSz As ADODB.Recordset
            Set ORSz = oRSdir.NextRecordset
            Set Me.DatZona.RowSource = ORSz
            Me.DatZona.BoundColumn = ORSz.Fields(0).Name
            Me.DatZona.ListField = ORSz.Fields(1).Name
            Me.DatZona.BoundText = -1
            
            'Me.DatDireccion.BoundText = orsP!IDZ
            Me.DatDireccion.BoundText = orsP!IDDIR
        End If
    End If

End Sub

Public Function AgregaPlato(vcp As Double, _
                            vc As Double, _
                            vpre As Double, _
                            vimp As Double, _
                            vd As String, _
                            vnumser As String, _
                            vNumFac As Double, _
                            VcLIENTE As String, _
                            VcOMENSALES As Integer, _
                            Optional ByRef vnumsec As Integer) As Boolean
    LimpiaParametros oCmdEjec

    Dim xPedido     As String

    Dim NumSer      As String

    Dim NumFac      As Double

    Dim vMaxNumoper As String

    On Error GoTo ErrorGraba

    oCmdEjec.CommandType = adCmdStoredProc
    Pub_ConnAdo.BeginTrans

    If VNuevo Then 'nueva comanda

        With oCmdEjec
            .CommandText = "SPREGISTRARCOMANDADELIVERY"
            .Parameters.Append .CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
            .Parameters.Append .CreateParameter("@Usuario", adVarChar, adParamInput, 10, LK_CODUSU)
            .Parameters.Append .CreateParameter("@CodMesa", adVarChar, adParamInput, 10, "00")
            .Parameters.Append .CreateParameter("@cp", adInteger, adParamInput, , vcp)
            .Parameters.Append .CreateParameter("@cant", adDouble, adParamInput, , vc)
            .Parameters.Append .CreateParameter("@pre", adDouble, adParamInput, , vpre)
            .Parameters.Append .CreateParameter("@imp", adDouble, adParamInput, , vimp)
            .Parameters.Append .CreateParameter("@d", adVarChar, adParamInput, 50, vd)

            '       .Parameters.Append .CreateParameter("@Total", adCurrency, adParamInput, , Val(Me.lblTot.Tag))
            .Parameters.Append .CreateParameter("@Mozo", adInteger, adParamInput, , 0)

            .Parameters.Append .CreateParameter("@NumSer", adChar, adParamOutput, 3, NumSer)
            .Parameters.Append .CreateParameter("@NumFac", adDouble, adParamOutput, , NumFac)
            .Parameters.Append .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
            .Parameters.Append .CreateParameter("@CodFam", adInteger, adParamInput, , vCodFam)

            'PARAMETROS NUEVOS
            
            .Parameters.Append .CreateParameter("@CLIENTE", adVarChar, adParamInput, 120, VcLIENTE) 'Linea nueva

            .Parameters.Append .CreateParameter("@COMENSALES", adDouble, adParamInput, , VcOMENSALES)  'Linea nueva
            .Parameters.Append .CreateParameter("@DIRECCION", adVarChar, adParamInput, 150, Me.DatDireccion.Text)
            .Parameters.Append .CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , Me.lblCliente.Caption)
             '.Parameters.Append .CreateParameter("@IDZONA", adInteger, adParamInput, , IIf(Me.DatDireccion.BoundText = "", -1, Me.DatDireccion.BoundText))
            .Parameters.Append .CreateParameter("@IDZONA", adInteger, adParamInput, , IIf(Me.DatZona.BoundText = "", -1, Me.DatZona.BoundText))
            .Parameters.Append .CreateParameter("@RECOJO", adBoolean, adParamInput, , Me.chkRecojo.Value)               '08-06-2020
            If Me.DatDireccion.BoundText = "" Then
            .Parameters.Append .CreateParameter("@IDDIR", adInteger, adParamInput, , -1)         '08-06-2020
            Else
            .Parameters.Append .CreateParameter("@IDDIR", adInteger, adParamInput, , Me.DatDireccion.BoundText)         '08-06-2020
            End If
           
            '=================
            .Parameters.Append .CreateParameter("@NumSec", adInteger, adParamOutput, , 0)
            .Execute
            ' , Array(LK_CODCIA, LK_CODUSU, vMesa, vcp, vc, vpre, vimp, "ss", CInt(Me.lblMozo.Tag), NumSer, NumFac, LK_FECHA_DIA, vCodFam, 0)
        
            Me.lblSerie.Caption = oCmdEjec.Parameters("@NumSer").Value
            Me.lblNumero.Caption = oCmdEjec.Parameters("@NumFac").Value
            vnumsec = oCmdEjec.Parameters("@NumSec").Value
        End With

    Else

        With oCmdEjec
            .CommandText = "SPMODIFICARCOMANDADELIVERY"
            .Parameters.Append .CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
            .Parameters.Append .CreateParameter("@Usuario", adVarChar, adParamInput, 10, LK_CODUSU)
            .Parameters.Append .CreateParameter("@CodMesa", adVarChar, adParamInput, 10, "00")
            .Parameters.Append .CreateParameter("@cp", adDouble, adParamInput, , vcp) 'julio 11-01-2011
            .Parameters.Append .CreateParameter("@cant", adDouble, adParamInput, , vc)
            .Parameters.Append .CreateParameter("@pre", adDouble, adParamInput, , vpre)
            .Parameters.Append .CreateParameter("@imp", adDouble, adParamInput, , vimp)
            .Parameters.Append .CreateParameter("@d", adVarChar, adParamInput, 50, vd)
        
            '       .Parameters.Append .CreateParameter("@Total", adCurrency, adParamInput, , Val(Me.lblTot.Tag))
            .Parameters.Append .CreateParameter("@Mozo", adInteger, adParamInput, , 0)
        
            .Parameters.Append .CreateParameter("@NumSer", adChar, adParamInput, 3, vnumser)
            .Parameters.Append .CreateParameter("@NumFac", adDouble, adParamInput, , vNumFac)
            .Parameters.Append .CreateParameter("@NUMSEC", adInteger, adParamOutput)
            .Parameters.Append .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
            .Parameters.Append .CreateParameter("@CodFam", adInteger, adParamInput, , vCodFam)  'linea nueva
            .Parameters.Append .CreateParameter("@CLIENTE", adVarChar, adParamInput, 120, VcLIENTE) 'Linea nueva

            .Parameters.Append .CreateParameter("@COMENSALES", adDouble, adParamInput, , VcOMENSALES)  'Linea nueva
            .Parameters.Append .CreateParameter("@DIRECCION", adVarChar, adParamInput, 150, Me.DatDireccion.Text)
            .Parameters.Append .CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , Me.lblCliente.Caption)
            '.Parameters.Append .CreateParameter("@ZONA", adInteger, adParamInput, , vCodZona)
            .Execute
            vnumsec = .Parameters("@NUMSEC").Value
        End With

    End If

    'oCmdEjec.Execute
    
    'actualiza stock
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpActualizaStock"

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Usuario", adVarChar, adParamInput, 20, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodArt", adDouble, adParamInput, , vcp)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@cp", adInteger, adParamInput, , 1)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@cp", adInteger, adParamInput, , vc)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ser", adChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@nro", adInteger, adParamInput, , Me.lblNumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@mesa", adVarChar, adParamInput, 10, "")
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@tipo", adBoolean, adParamInput, , 1) '0 cuando es extorno
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumSec", adInteger, adParamInput, , vnumsec)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MaxNumFac", adDouble, adParamOutput, , vMaxFac)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MaxNumOper", adDouble, adParamOutput, , 0)

    oCmdEjec.Execute
    vMaxFac = oCmdEjec.Parameters("@MaxNumFac").Value
    vMaxNumoper = oCmdEjec.Parameters("@MaxNumOper").Value

    AgregaPlato = True
    Pub_ConnAdo.CommitTrans

    Exit Function

ErrorGraba:
    AgregaPlato = False
    Pub_ConnAdo.RollbackTrans
    MsgBox Err.Description
End Function

Private Sub ConfiguraLV()

    With Me.lvPlatos
        .Gridlines = True
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .View = lvwReport
    
        .HideSelection = False
        .ColumnHeaders.Add , , "Descripción", 5200
        .ColumnHeaders.Add , , "Cta.", 0
        .ColumnHeaders.Add , , "Detalle", 1000
        .ColumnHeaders.Add , , "Cant.", 900, 1
        .ColumnHeaders.Add , , "Precio", 1000, 1
        .ColumnHeaders.Add , , "Importe", 1200, 1
        .ColumnHeaders.Add , , "numsec", 0
        .ColumnHeaders.Add , , "Atend", 0
        .ColumnHeaders.Add , , "allnumfac", 0
        .ColumnHeaders.Add , , "apro", 0
        .ColumnHeaders.Add , , "numfac", 0
        .MultiSelect = True
    End With

    With Me.lvCliente
        .Gridlines = True
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .View = lvwReport
    
        .ColumnHeaders.Add , , "Cliente", 5000
        .ColumnHeaders.Add , , "DNI", 1000
    End With

    With Me.ListView1
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .View = lvwReport
  
        .ColumnHeaders.Add , , "Cliente", 5000
  
        .MultiSelect = False
    End With

End Sub

Private Sub CargarPlatos()
oCmdEjec.CommandText = "SpListarPlatosDelivery"
oCmdEjec.CommandType = adCmdStoredProc
Set oRsPlatos = oCmdEjec.Execute(, Array(LK_CODCIA))
End Sub

Private Sub CargarSubFamilias()
oCmdEjec.CommandText = "SpListarSubFamilias"
oCmdEjec.CommandType = adCmdStoredProc
Set oRsSubFam = oCmdEjec.Execute(, LK_CODCIA)
End Sub

Private Sub FiltrarPlatos(cant As Integer, oRS As ADODB.Recordset)
vPlato = cant
'Dim vPri As Boolean
'vPri = True
Dim f, c As Integer
c = 1

Dim valor As Double
valor = vPlato / 22
' agregado julio 06/08/2012==============
 pos = InStr(Trim(str(valor)), ".")
    If pos <> 0 Then
        If pos = 1 Then
            ent = Left(CStr(valor), pos)
        Else
            ent = Left(CStr(valor), pos - 1)
        End If

    Else
        ent = Int(valor)
    End If
    
    If pos <> 0 Then
        pos2 = Right(Trim(str(valor)), Len(Trim(str(valor))) - pos)
    Else
        pos2 = 0
    End If

    If pos2 > 0 Then
        ent = ent + 1
    End If

    vPagTotPla = ent

    If valor <> 0 Then vPagActPla = 1
' agregado julio 06/08/2012==============

' quitado julio 06/08/2012==============
''pos = InStr(Trim(str(valor)), ".")
''If pos > 0 Then
''pos2 = Right(Trim(str(valor)), Len(Trim(str(valor))) - pos)
''ENT = Left(Trim(str(valor)), pos - 1)
''If ENT = "" Then: ENT = 0
''If pos2 > 0 Then: vPagTotPla = ENT + 1
''End If
''If vPlato >= 1 Then: vPagActPla = 1
' quitado julio 06/08/2012==============


If vPlato > 18 Then: Me.cmdPlatoSig.Enabled = True
'descargar los objetos primero
If Me.cmdPlato.count > 1 Then
    For i = 1 To cmdPlato.count - 1
        Unload cmdPlato.Item(i)
    Next
End If
'============================
vIniLeft = 30
vIniTop = 120
For i = 1 To vPlato
    Load Me.cmdPlato(i)
    If c <= 5 Then '1 fila
        If c = 1 Then
            vIniLeft = vIniLeft + Me.cmdPlatoAnt.Width
        Else
            vIniLeft = vIniLeft + Me.cmdPlato(i - i).Width
        End If
'        Else: viniLeft = viniLeft + 970
'        End If
    ElseIf c <= 11 Then '2º Fila
        'viniLeft = 30
        If c = 6 Then
            vIniLeft = 30
            vIniTop = vIniTop + Me.cmdPlatoAnt.Height
        Else: vIniLeft = vIniLeft + Me.cmdPlato(i - 1).Width
        End If
    ElseIf c <= 17 Then '3º Fila
        If c = 12 Then
            vIniTop = vIniTop + Me.cmdPlato(4).Height
            vIniLeft = 30
        Else: vIniLeft = vIniLeft + Me.cmdPlato(i - 1).Width
        End If
    Else '4º y ultima fila
        If c = 18 Then
            vIniTop = vIniTop + Me.cmdPlato(11).Height
            vIniLeft = 30
        Else: vIniLeft = vIniLeft + Me.cmdPlato(i - 1).Width
        End If
    End If
    Me.cmdPlato(i).Left = vIniLeft
    Me.cmdPlato(i).Top = vIniTop
    Me.cmdPlato(i).Visible = True
    'Me.cmdPlato(i).BackColor = Me.cmdSubFam(vColor).BackColor   'gts para mostrar el color de la familia
    'CARGANDO LA IMAGEN
    
     If Not IsNull(oRS!Datoimagen) Then
     
        Dim sIMG As ADODB.Stream
        
        ' Nuevo objeto Stream para poder leer el campo de imagen
        Set sIMG = New ADODB.Stream
        
        ' Especifica el tipo de datos ( binario )
        sIMG.Type = adTypeBinary
        sIMG.Open
        ' Graba los datos en el objeto stream
        sIMG.Write oRS.Fields!Datoimagen
        
        ' este método graba un  archivo temporal  en disco _
        ( en el app.path que luego se elimina )
        sIMG.SaveToFile App.Path & "\temp.jpg", adSaveCreateOverWrite
        
        'AGREGA LA IMAGEN
        
        Me.cmdPlato(i).Picture = LoadPicture(App.Path & "\temp.jpg")
        ' Elimina el archivo temporal
        Kill App.Path & "\temp.jpg"
        
        
        If sIMG.State = adStateOpen Then sIMG.Close
        If Not sIMG Is Nothing Then Set sIMG = Nothing
    End If
    '==================
    
    'Me.cmdPlato(i).Style = 1
    'Me.cmdFam(i).Visible = vPri
    Me.cmdPlato(i).Visible = True
    Me.cmdPlato(i).Caption = Trim(oRS!plato)
    Me.cmdPlato(i).ToolTipText = Trim(oRS!alt)
    Me.cmdPlato(i).Tag = oRS!Codigo
  '  Me.cmdPlato(i).BackColor = Trim(ORS!alt)
'    If c <= 14 Then
'        Me.cmdFam(i).Visible = True
'    Else
'        Me.cmdFam(i).Visible = False
'    End If
oRS.MoveNext
    If c = 22 Then
'        vPri = False
        c = 1
        'vuelve a empezar
        vIniLeft = 30
        vIniTop = 120
        Else
        c = c + 1
   End If
   
Next
End Sub


Private Sub FiltarSubFamilias(cant As Integer, oRS As ADODB.Recordset)

vSubFamilia = cant
'Dim vPri As Boolean
'vPri = True
Dim f, c As Integer


c = 1

Dim valor As Double
valor = vSubFamilia / 14

pos = InStr(Trim(str(valor)), ".")
pos2 = Right(Trim(str(valor)), Len(Trim(str(valor))) - pos)
If pos > 0 Then ent = Left(Trim(str(valor)), pos - 1) 'JULIO 09/01/2011
'ent = Left(Trim(str(valor)), pos - 1)
If ent = "" Then: ent = 0
If pos2 > 0 Then: vPagTotSubFam = ent + 1

If vSubFamilia >= 1 Then: vPagActSubFam = 1
If vSubFamilia > 14 Then: Me.cmdSubFamSig.Enabled = True
'descargar los objetos primero
If cmdSubFam.count > 1 Then
    For i = 1 To cmdSubFam.count - 1
        Unload cmdSubFam.Item(i)
    Next
End If
'============================
vIniLeft = 30
vIniTop = 120
For i = 1 To vSubFamilia
    Load Me.cmdSubFam(i)
    
    If c <= 3 Then '1 fila
    
        vIniLeft = vIniLeft + Me.cmdSubFamAnt.Width
        Me.cmdSubFam(i).Left = vIniLeft
        Me.cmdSubFam(i).Top = vIniTop
        Me.cmdSubFam(i).Visible = True
        
    ElseIf c <= 7 Then '2º Fila
        'viniLeft = 30
        If c = 4 Then
            vIniLeft = 30
            vIniTop = vIniTop + Me.cmdSubFamAnt.Height
        Else: vIniLeft = vIniLeft + Me.cmdSubFamAnt.Width
        End If
    ElseIf c <= 11 Then '3º Fila
        If c = 8 Then
            vIniTop = vIniTop + Me.cmdSubFam(4).Height
            vIniLeft = 30
        Else: vIniLeft = vIniLeft + Me.cmdSubFamAnt.Width
        End If
    Else '4º y ultima fila
        If c = 12 Then
            vIniTop = vIniTop + Me.cmdSubFam(11).Height
            vIniLeft = 30
        Else: vIniLeft = vIniLeft + Me.cmdSubFamAnt.Width
        End If
    End If
    Me.cmdSubFam(i).Left = vIniLeft
    Me.cmdSubFam(i).Top = vIniTop
    'Me.cmdFam(i).Visible = vPri
    Me.cmdSubFam(i).Visible = True
    Me.cmdSubFam(i).Caption = Trim(oRS!Familia)
    Me.cmdSubFam(i).Tag = oRS!NUMERO
 'If Not IsNull(ors!color) Then Me.cmdSubFam(i).BackColor = Trim(ors!color)  GTS ACA QUITO EL COLOR A FAMILIAS
    
'    If c <= 14 Then
'        Me.cmdFam(i).Visible = True
'    Else
'        Me.cmdFam(i).Visible = False
'    End If
oRS.MoveNext
    If c = 14 Then
'        vPri = False
        c = 1
        'vuelve a empezar
        vIniLeft = 30
        vIniTop = 120
        Else
        c = c + 1
   End If
   
Next


    

End Sub

Private Sub chkRecojo_Click()
If Me.chkRecojo.Value = 0 Then
    Me.chkRecojo.Caption = "Delivery"
Else
Me.chkRecojo.Caption = "Recojo en Tienda"
End If

    If gEstaCargando Then Exit Sub
    If VNuevo = False Then ' EDITANDO COMANDA

        On Error GoTo OBS

        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "USP_PEDIDO_RECOJO"
        oCmdEjec.CommandType = adCmdStoredProc
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adDouble, adParamInput, , Me.lblSerie.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSFAC", adDouble, adParamInput, , Me.lblNumero.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RECOJO", adBoolean, adParamInput, , Me.chkRecojo.Value)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adInteger, adParamInput, , IIf(Me.DatDireccion.BoundText = "", -1, Me.DatDireccion.BoundText))

        oCmdEjec.Execute
        
        oCmdEjec.CommandText = "USP_COMANDA_ACTUALIZA_PAGOSDELIVERY"
        LimpiaParametros oCmdEjec
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adVarChar, adParamInput, 3, Me.lblSerie.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adDouble, adParamInput, , Me.lblNumero.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TOTAL", adDouble, adParamInput, , Me.lblTot.Caption)
        'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adDouble, adParamInput, , Me.DatZona.BoundText)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adDouble, adParamInput, , Me.DatZona.BoundText)
    
        oCmdEjec.Execute
    
        MsgBox "Datos Almacenados Correctamente.", vbInformation, Pub_Titulo

        Exit Sub

OBS:
        MsgBox Err.Description, vbCritical, Pub_Titulo

    End If
End Sub

Private Sub cmdActualizaObservaciones_Click()

    If Me.lvPlatos.ListItems.count = 0 Then
        MsgBox "Debe agregar un articulo para actualizar observaciones", vbInformation, Pub_Titulo

        Exit Sub

    End If

    On Error GoTo OBS

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_PEDIDO_OBSERVACION"
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adDouble, adParamInput, , Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSFAC", adDouble, adParamInput, , Me.lblNumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)

    If Len(Trim(Me.txtObs.Text)) <> 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@OBS", adVarChar, adParamInput, 200, Me.txtObs.Text)
    End If

    oCmdEjec.Execute
    MsgBox "Datos Almacenados Correctamente.", vbInformation, Pub_Titulo

    Exit Sub

OBS:
    MsgBox Err.Description, vbCritical, Pub_Titulo
End Sub

Private Sub cmdBorrar_Click()
If Len(Me.lblTexto.Caption) > 0 Then
    Me.lblTexto.Caption = Left(Me.lblTexto.Caption, Len(Me.lblTexto.Caption) - 1)
End If
End Sub

Private Sub cmdCantidad_Click()

    On Error GoTo elimina

    If Not Me.lvPlatos.SelectedItem Is Nothing Then

        ' If Me.lvPlatos.SelectedItem.SubItems(6) <> 1 Or Me.lvPlatos.SelectedItem.SubItems(5) <> 0 Then
        'If Me.lvPlatos.SelectedItem.SubItems(9) = 1 Then 'Or val(Me.lvPlatos.SelectedItem.SubItems(7)) <> 0 Then
        If Me.lvPlatos.SelectedItem.SubItems(9) = 1 Or Me.lvPlatos.SelectedItem.SubItems(7) <> 0 Or vEstado = "E" Then
            'If MsgBox("no se puede MODIFICAR el plato, ya fue despachado" & vbCrLf & "¿desea ingresar la clave?", vbQuestion + vbYesNo, NombreProyecto) = vbYes Then
            MsgBox ("no se puede MODIFICAR el plato, ya fue despachado")
           
        Else
    Me.lblicbper.Caption = val(Me.lblicbperORI.Caption) * val(Me.lblTexto.Caption)
            'Varificando insumos del plato
            Dim msn As String

            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SpDevuelveInsumosxPlato"
            oCmdEjec.CommandType = adCmdStoredProc
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodPlato", adDouble, adParamInput, , Me.lvPlatos.SelectedItem.Tag)
            'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@mensaje", adVarChar, adParamOutput, 300, msn)
 
            Dim vstrmin  As String 'variable para capturar los insumos

            Dim vstrcero As String 'variabla para capturar insumos en cero

            Dim vmin     As Boolean 'minimo

            Dim vcero    As Boolean 'stock cero

            vmin = False
            vcero = False

            Dim ss       As Currency

            Dim stockReq As Double
 
            Set oRStemp = oCmdEjec.Execute

            If Not oRStemp.EOF Then
 
                ss = val(Me.lblTexto.Caption) - val(Me.lvPlatos.SelectedItem.SubItems(3))

                If ss = 0 Then Exit Sub
                If ss > 0 Then
    
                    Do While Not oRStemp.EOF
        
                        stockReq = val(ss) * oRStemp!ei
        
                        If oRStemp!sa <= 0 Or stockReq > oRStemp!sa Then
                            vcero = True    'LINEA FALTANTE

                            If Len(vstrcero) = 0 Then
                                vstrcero = Trim(oRStemp!nm)
                            Else
                                vstrcero = vstrcero & vbCrLf & Trim(oRStemp!nm)
                            End If

                        ElseIf (val(oRStemp!sa) - val(stockReq)) <= oRStemp!sm Then
                            vmin = True    'LINEA FALTANTE

                            If Len(vstrmin) = 0 Then
                                vstrmin = Trim(oRStemp!nm)
                            Else
                                vstrmin = vstrmin & vbCrLf & Trim(oRStemp!nm)
                            End If
                        End If

                        oRStemp.MoveNext
       
                        'c = c + 1
        
                    Loop
   
                End If

            End If
 
            If vmin Then
                MsgBox "Los siguientes insumos del Plato estan el el Minimó permitido" & vbCrLf & vstrmin, vbInformation, NombreProyecto
            End If
 
            If vcero Then
                'MsgBox "Algunos insumos del Plato no estan disponibles" & vbCrLf & vstrcero, vbCritical, NombreProyecto
            End If

            'If vcero Then Exit Sub
 
            If Len(Trim(Me.lblTexto.Caption)) <> 0 Then
                If Not Me.lvPlatos.SelectedItem Is Nothing Then
                    LimpiaParametros oCmdEjec

                    With oCmdEjec
                        .CommandText = "SpModificarPreCantPla"
                        .Parameters.Append .CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                        .Parameters.Append .CreateParameter("@NumSer", adChar, adParamInput, 4, Trim(Me.lblSerie.Caption))
                        .Parameters.Append .CreateParameter("@NumFac", adDouble, adParamInput, , CDbl(Me.lblNumero.Caption))
                        .Parameters.Append .CreateParameter("@NumSec", adInteger, adParamInput, , CInt(Me.lvPlatos.SelectedItem.SubItems(6)))
                        .Parameters.Append .CreateParameter("@CodArt", adDouble, adParamInput, , CDbl(Me.lvPlatos.SelectedItem.Tag))
                        .Parameters.Append .CreateParameter("@Pre", adInteger, adParamInput, , Null)
                        .Parameters.Append .CreateParameter("@Cant", adDouble, adParamInput, , CDbl(Me.lblTexto.Caption))
                        .Parameters.Append .CreateParameter("@EsPre", adBoolean, adParamInput, , False)
                        .Execute
        
                        'actualiza stock
                        LimpiaParametros oCmdEjec
                        .CommandText = "SpActualizarStockComanda"
        
                        .Parameters.Append .CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                        .Parameters.Append .CreateParameter("@codplato", adDouble, adParamInput, , Me.lvPlatos.SelectedItem.Tag)
                        .Parameters.Append .CreateParameter("@nf", adInteger, adParamInput, , Me.lvPlatos.SelectedItem.SubItems(10))
                        '.Parameters.Append .CreateParameter("@nf", adInteger, adParamInput, , Me.lvPlatos - .S)
                        .Parameters.Append .CreateParameter("@cant", adInteger, adParamInput, 3, Me.lblTexto.Caption)
                        .Parameters.Append .CreateParameter("@sa", adDouble, adParamInput, , Me.lvPlatos.SelectedItem.SubItems(3))
                        .Parameters.Append .CreateParameter("@fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA) 'JULIO 13/02/2011
        
                        oCmdEjec.Execute
 
                    End With

                    Me.lvPlatos.SelectedItem.SubItems(3) = Format(Me.lblTexto.Caption, "##.#0")
                    Me.lvPlatos.SelectedItem.SubItems(5) = Format(CStr(val(Me.lvPlatos.SelectedItem.SubItems(3)) * val(Me.lvPlatos.SelectedItem.SubItems(4))), "##.#0")
                    Me.lblTot.Caption = FormatCurrency(sumatoria, 2)
                    Me.lblTexto.Caption = ""
                End If
            End If
        End If
    End If

    oCmdEjec.CommandText = "USP_COMANDA_ACTUALIZA_PAGOSDELIVERY"
    LimpiaParametros oCmdEjec
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adVarChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adDouble, adParamInput, , Me.lblNumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TOTAL", adDouble, adParamInput, , Me.lblTot.Caption)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adDouble, adParamInput, , Me.DatZona.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adDouble, adParamInput, , Me.DatZona.BoundText)
    
    oCmdEjec.Execute



    Exit Sub

elimina:
    MsgBox Err.Description

End Sub

Private Sub cmdCaracteristicas_Click()
If Me.lvPlatos.ListItems.count = 0 Then Exit Sub
frmComandaProdCaracteristicas.gIDproducto = Me.lvPlatos.SelectedItem.Tag
frmComandaProdCaracteristicas.gNUMFAC = Me.lblNumero.Caption
frmComandaProdCaracteristicas.gNUMSER = Me.lblSerie.Caption
frmComandaProdCaracteristicas.gNUMSEC = Me.lvPlatos.SelectedItem.SubItems(6)
frmComandaProdCaracteristicas.Show vbModal
End Sub

Private Sub cmdCliente_Click()

    If Len(Trim(Me.lblCliente.Caption)) = 0 Then
        MsgBox "Debe elegir el cliente primero.", vbInformation, Pub_Titulo

        Exit Sub

    End If

    frmClientesDir.gIDcliente = Me.lblCliente.Caption
    frmClientesDir.gCliente = Me.txtCliente.Text

    'frmDeliveryClienteDireccion.Show vbModal

    frmClientesDir.Show vbModal

    If frmClientesDir.gGraba Then
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SP_DELIVERY_CLIENTE_DIRECCIONES"
        Me.lblReferencia.Caption = ""
        Set oRSdir = Nothing
        Set oRSdir = oCmdEjec.Execute(, Array(LK_CODCIA, Me.lblCliente.Caption))
        Set Me.DatDireccion.RowSource = Nothing
        Me.DatDireccion.BoundText = ""
        Set Me.DatDireccion.RowSource = oRSdir
        Me.DatDireccion.BoundColumn = oRSdir.Fields(4).Name
        Me.DatDireccion.ListField = oRSdir.Fields(0).Name
        Me.DatDireccion.SetFocus
            
        'Dim ORSurb As ADODB.Recordset
        Set ORSurb = oRSdir.NextRecordset
        Me.txtUrb.Text = ORSurb!urb
        Me.lblurb.Caption = ORSurb!IDEURB
            
        Dim ORSz As ADODB.Recordset

        Set ORSz = oRSdir.NextRecordset
        Set Me.DatZona.RowSource = ORSz
        Me.DatZona.BoundColumn = ORSz.Fields(0).Name
        Me.DatZona.ListField = ORSz.Fields(1).Name
        Me.DatZona.BoundText = -1
    End If

End Sub

Private Sub cmdClienteEdit_Click()

    If Me.cmdClienteEdit.Caption = "Cambiar Cliente" Then
        Me.FraCliente.Enabled = True
        Me.txtCliente.SetFocus
        Me.txtCliente.SelStart = 0
        Me.txtCliente.SelLength = Len(Me.txtCliente.Text)
        
        Me.cmdClienteEdit.Caption = "Grabar Cliente"
    Else
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SP_DELIVERY_CAMBIARCLIENTE"

        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adInteger, adParamInput, , Me.lblSerie.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adBigInt, adParamInput, , Me.lblNumero.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , Me.lblCliente.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DIRECCION", adVarChar, adParamInput, 150, Me.DatDireccion.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adBigInt, adParamInput, , Me.DatZona.BoundText)
        
        If Me.DatDireccion.BoundText = "" Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDDIR", adBigInt, adParamInput, , -1)
        Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDDIR", adBigInt, adParamInput, , Me.DatDireccion.BoundText)
        End If
        
        Dim VEXITO As String

        VEXITO = ""
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@EXITO", adVarChar, adParamOutput, 300, VEXITO)
        oCmdEjec.Execute

        If Len(Trim(VEXITO)) <> 0 Then
            MsgBox VEXITO, vbCritical, Pub_Titulo
        Else
            MsgBox "Datos Almacenados Correctamente."
        End If

        Me.FraCliente.Enabled = False
        Me.cmdClienteEdit.Caption = "Cambiar Cliente"
    End If

End Sub

Private Sub cmdconsulta_Click()

    If Len(Trim(Me.lblCliente.Caption)) = 0 Then
        MsgBox "Debe elegir el Cliente.", vbInformation, Pub_Titulo
        Me.txtCliente.SetFocus
        Me.txtCliente.SelStart = 0
        Me.txtCliente.SelLength = Len(Me.txtCliente.Text)
        Exit Sub
    End If

    frmDeliveryUltCompras.gICLIENTE = Me.lblCliente.Caption
    frmDeliveryUltCompras.lblCliente.Caption = Me.txtCliente.Text
    frmDeliveryUltCompras.Show vbModal
End Sub

Private Sub cmdDescuentos_Click()
    frmClaveCaja.Show vbModal

    If frmClaveCaja.vAceptar Then
        Dim vS As String
        If VerificaPass(frmClaveCaja.vUSUARIO, frmClaveCaja.vClave, vS) Then
            frmComandaDescuentos.Show vbModal

            If frmComandaDescuentos.gAcepta Then
                LimpiaParametros oCmdEjec
                oCmdEjec.CommandText = "SP_CALCULAR_DESCUENTOS"
        
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adChar, adParamInput, 3, Me.lblSerie.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adDouble, adParamInput, , CDbl(Me.lblNumero.Caption))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 20, frmClaveCaja.vUSUARIO)
        
                If frmComandaDescuentos.gTIPO = "2" Then '"P" Then 'porcentual
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DESCUENTO", adDouble, adParamInput, , frmComandaDescuentos.gDSCTO)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPO", adInteger, adParamInput, , 2)
                ElseIf frmComandaDescuentos.gTIPO = "1" Then 'total
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DESCUENTO", adDouble, adParamInput, , frmComandaDescuentos.gDSCTO)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPO", adInteger, adParamInput, , 1)
                Else 'por familia
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DESCUENTO", adDouble, adParamInput, , frmComandaDescuentos.gDSCTO)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPO", adInteger, adParamInput, , 3)
                End If

                oCmdEjec.Execute

              'CargarComanda LK_CODCIA, vMesa
               CargarComanda LK_CODCIA, vnumser, vNumFac
            End If

        Else
            MsgBox "Clave incorrecta", vbCritical, Pub_Titulo
        End If
    End If

End Sub

Private Sub cmdDetalle_Click()

    If Not Me.lvPlatos.SelectedItem Is Nothing Then

        Dim Dato As String
        frmDetalle.EsDetalle = True
        frmDetalle.Show vbModal
       
 frmDetalle.Comensales = False
        If frmDetalle.vSelec Then

            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SpActualizarDetallePlato"
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adChar, adParamInput, 3, Me.lblSerie.Caption)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adDouble, adParamInput, , CDbl(Me.lblNumero.Caption))
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numsec", adInteger, adParamInput, , CInt(Me.lvPlatos.SelectedItem.SubItems(6)))
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@det", adVarChar, adParamInput, 50, frmDetalle.vDetalle)
            oCmdEjec.Execute
            Me.lvPlatos.SelectedItem.SubItems(2) = frmDetalle.vDetalle
        End If
    End If


End Sub

Private Sub cmdDoctoDespacho_Click()

    If Me.lvPlatos.ListItems.count = 0 Then
        MsgBox "Debe agregar platos al Delivery", vbInformation, Pub_Titulo

        Exit Sub

    End If

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPVALIDAPEDIDOFACTURADO"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim rsd As ADODB.Recordset

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumSer", adInteger, adParamInput, , Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumFac", adDouble, adParamInput, , Me.lblNumero.Caption)

    Set rsd = oCmdEjec.Execute
    
    If Not rsd.EOF Then
        If Not CBool(rsd!fac) Then
            MsgBox "Debe facturar el Delivery antes de imprimir el Documento de Despacho.", vbInformation, Pub_Titulo

            Exit Sub

        End If
    End If

    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions

    Dim crParamDef  As CRAXDRT.ParameterFieldDefinition

    Dim objCrystal  As New CRAXDRT.APPLICATION

    Dim RutaReporte As String

    RutaReporte = "C:\Admin\Nordi\DoctoDespacho.rpt"

    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
    Set crParamDefs = VReporte.ParameterFields

    For Each crParamDef In crParamDefs

        Select Case crParamDef.ParameterFieldName

            Case "mesa"
                crParamDef.AddCurrentValue str(vPlato)
        End Select

    Next

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.CommandText = "SPPRINTDESPACHO"
    'oCmdEjec.CommandText = "SpPrintComanda"
 
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumSer", adChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumFac", adDouble, adParamInput, , Me.lblNumero.Caption)

    Set rsd = oCmdEjec.Execute
    
      'SUB REPORTE
    Dim VReporteS As New CRAXDRT.Report
    

    VReporte.DataBase.SetDataSource rsd, , 1

    Set VReporteS = VReporte.OpenSubreport("PAGOS")
    
     Set crParamDefs = VReporteS.ParameterFields

    For Each crParamDef In crParamDefs

        Select Case crParamDef.ParameterFieldName

            Case "Pm-ado.NROCOMANDA"
                crParamDef.AddCurrentValue Me.lblSerie.Caption + "-" + Me.lblNumero.Caption
        End Select

    Next
    
   
    
    VReporte.OpenSubreport("PAGOS").DataBase.LogOnServer "p2sodbc.dll", "DSN_DATOS", "bdatos", "sa", claveAcceso
    
    VReporte.PrintOut False, 1, , 1, 1
    Set objCrystal = Nothing
    Set VReporte = Nothing
End Sub

Private Sub cmdEliminar_Click()

    On Error GoTo elimina

    If Not Me.lvPlatos.SelectedItem Is Nothing Then
        If Me.lvPlatos.SelectedItem.SubItems(9) = 1 Or Me.lvPlatos.SelectedItem.SubItems(7) <> 0 Or vEstado = "E" Then  'Or val(Me.lvPlatos.SelectedItem.SubItems(6)) <> 0 Then
            If MsgBox("no se puede eliminar el plato, ya fue despachado o la Mesa está EN CUENTA" & vbCrLf & "¿desea ingresar la clave?", vbQuestion + vbYesNo, NombreProyecto) = vbYes Then
                frmClaveCaja.Show vbModal

                If frmClaveCaja.vAceptar Then

                    Dim vS As String

                    If VerificaPass(frmClaveCaja.vUSUARIO, frmClaveCaja.vClave, vS) Then
                        frmComandaMotivoElimina.Show vbModal

                        If frmComandaMotivoElimina.gAcepta Then
                            extorna frmComandaMotivoElimina.gIDmotivo, frmComandaMotivoElimina.gMOTIVO, frmClaveCaja.vUSUARIO
                        End If
                           'extorna
                    Else
                        MsgBox "Clave incorrecta", vbCritical, NombreProyecto
                    End If
                End If
            End If

        Else
            elimina
        End If
    End If

    Exit Sub

elimina:
    MsgBox Err.Description

End Sub

Private Sub cmdEnviar_Click()

    If Len(Trim(Me.lblNumero.Caption)) = 0 And Len(Trim(Me.lblNumero.Caption)) = 0 Then
        MsgBox "No hay nada que imprimir", vbCritical, NombreProyecto
    Else
        Me.cmdClienteEdit.Visible = False
        Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions

        Dim crParamDef  As CRAXDRT.ParameterFieldDefinition

        Dim objCrystal  As New CRAXDRT.APPLICATION

        Dim RutaReporte As String

        RutaReporte = "C:\Admin\Nordi\ComandaDelivery.rpt"
        'RutaReporte = "C:\Admin\Nordi\Comanda.rpt"

        'Verificar platos enviados para mensaje
        Dim cat      As Integer

        Dim Mensaje  As String

        Dim mATRIZ() As Integer

        Dim ss       As Integer

        For cat = 1 To Me.lvPlatos.ListItems.count

            If Me.lvPlatos.ListItems(cat).Checked Then
                ReDim Preserve mATRIZ(ss)
                mATRIZ(ss) = cat
                ss = ss + 1
            End If

        Next

        Dim OVARIAN As Variant

        If ss > 0 Then

            For Each OVARIAN In mATRIZ

                If Me.lvPlatos.ListItems(OVARIAN).SubItems(9) = "1" Then
                    Mensaje = "DUPLICADO"
                Else
                    Mensaje = ""

                    Exit For

                End If

            Next

        End If

        Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
        Set crParamDefs = VReporte.ParameterFields

        For Each crParamDef In crParamDefs

            Select Case crParamDef.ParameterFieldName

                Case "mesa"
                    crParamDef.AddCurrentValue "" 'str(vPlato)

                Case "Mensaje"
                    crParamDef.AddCurrentValue Mensaje

                Case "pCLIENTE"
                    crParamDef.AddCurrentValue Me.txtCliente.Text
                    
                Case "pDESPACHO"
                    crParamDef.AddCurrentValue UCase(Me.chkRecojo.Caption)
                        
            End Select

        Next

        On Error GoTo printe

        LimpiaParametros oCmdEjec
        oCmdEjec.CommandType = adCmdStoredProc
        oCmdEjec.CommandText = "SP_DELIVERY_PRINT"
        'oCmdEjec.CommandText = "SpPrintComanda"

        Dim rsd     As ADODB.Recordset

        Dim vdata   As String

        Dim vnumsec As String

        vdata = ""

        Dim c As Integer

        If Me.lvPlatos.CheckBoxes Then

            For c = 1 To Me.lvPlatos.ListItems.count

                If Me.lvPlatos.ListItems(c).Checked Then
                    vdata = vdata & Me.lvPlatos.ListItems(c).Tag & ","
                    vnumsec = vnumsec & Me.lvPlatos.ListItems(c).SubItems(6) & ","
                End If

            Next

        Else

            For c = 1 To Me.lvPlatos.ListItems.count

                vdata = vdata & Me.lvPlatos.ListItems(c).Tag & ","
                vnumsec = vnumsec & Me.lvPlatos.ListItems(c).SubItems(6) & ","

            Next

        End If

        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumSer", adChar, adParamInput, 3, Me.lblSerie.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumFac", adDouble, adParamInput, , Me.lblNumero.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@xdet", adVarChar, adParamInput, 4000, vdata)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@xnumsec", adVarChar, adParamInput, 4000, vnumsec)

        Set rsd = oCmdEjec.Execute

        'OBTENER LAS FAMILIAS DE LA TABLA TABLAS

        Dim ORSf As ADODB.Recordset

        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SP_FAMILIAS_LISTPRINT"
        Set ORSf = oCmdEjec.Execute(, LK_CODCIA)

        Dim sFILTRO As String

        Dim oRStmp  As ADODB.Recordset

        Set oRStmp = New ADODB.Recordset
        oRStmp.CursorType = adOpenDynamic ' setting cursor type
        oRStmp.Fields.Append "FAMILIA", adVarChar, 100
        'oRSfp.Fields.Append "formapago", adVarChar, 120

        oRStmp.Fields.Refresh
        oRStmp.Open

        Dim MyMatriz() As String

        Do While Not ORSf.EOF
            MyMatriz = Split(ORSf!Familia, "|")

            For i = LBound(MyMatriz) To UBound(MyMatriz)

                'Le asignamos unos elementos de prueba
                If MyMatriz(i) <> "" Then
                    oRStmp.AddNew
                    oRStmp!Familia = MyMatriz(i)
                    oRStmp.Update
                End If

            Next

            sFILTRO = ""

            Dim IC As Integer

            If oRStmp.RecordCount <> 0 Then oRStmp.MoveFirst
            IC = 1

            Do While Not oRStmp.EOF

                If IC = oRStmp.RecordCount Then
                    sFILTRO = sFILTRO & "PED_FAMILIA=" & oRStmp!Familia
                Else
                    sFILTRO = sFILTRO & "PED_FAMILIA=" & oRStmp!Familia & " OR "
                End If

                IC = IC + 1
                oRStmp.MoveNext
            Loop

            rsd.Filter = sFILTRO

            If Not rsd.EOF Then
                VReporte.DataBase.SetDataSource rsd, 3, 1 'lleno el objeto reporte

                VReporte.SelectPrinter Printer.DriverName, ORSf!IMPRESORA, Printer.Port

                VReporte.PrintOut False, 1, , 1, 1

                Set VReporte = Nothing
                Set VReporte = objCrystal.OpenReport(RutaReporte, 1)

                Set crParamDefs = VReporte.ParameterFields

                For Each crParamDef In crParamDefs

                    Select Case crParamDef.ParameterFieldName

                        Case "mesa"
                            crParamDef.AddCurrentValue str(vPlato)

                        Case "Mensaje"
                            crParamDef.AddCurrentValue Mensaje
                            
                        Case "pCLIENTE"
                    crParamDef.AddCurrentValue Me.txtCliente.Text
                    
                    Case "pDESPACHO"
                    crParamDef.AddCurrentValue UCase(Me.chkRecojo.Caption)
                            
                    End Select

                Next
               If ORSf!IMPRESORA2 <> "" Then
                    VReporte.DataBase.SetDataSource rsd, 3, 1 'lleno el objeto reporte
                
                    VReporte.SelectPrinter Printer.DriverName, ORSf!IMPRESORA2, Printer.Port
                
                    VReporte.PrintOut False, 1, , 1, 1

                    Set VReporte = Nothing
                    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)

                    Set crParamDefs = VReporte.ParameterFields

                    For Each crParamDef In crParamDefs

                        Select Case crParamDef.ParameterFieldName

                            Case "mesa"
                                crParamDef.AddCurrentValue str(vPlato)

                            Case "Mensaje"
                                crParamDef.AddCurrentValue Mensaje
                            
                            Case "pCLIENTE"
                    crParamDef.AddCurrentValue Me.txtCliente.Text
                    
                    Case "pDESPACHO"
                    crParamDef.AddCurrentValue UCase(Me.chkRecojo.Caption)
                                
                        End Select

                    Next
 
            End If

            End If

            If Not oRStmp Is Nothing Then

                'If Not oRSfp.EOF Then oRSfp.Delete
                If oRStmp.RecordCount <> 0 Then
                    oRStmp.MoveFirst

                    Do While Not oRStmp.EOF
                        oRStmp.Delete adAffectCurrent
                        oRStmp.MoveNext
                    Loop

                End If
            End If

            ORSf.MoveNext
        Loop

        Set objCrystal = Nothing
        Set VReporte = Nothing

        Dim ct As Integer

        For ct = 1 To Me.lvPlatos.ListItems.count

            If Me.lvPlatos.ListItems(ct).Checked Then
                Me.lvPlatos.ListItems(ct).SubItems(9) = "1"
                Me.lvPlatos.ListItems(ct).Checked = False
            End If

        Next

        Me.lvPlatos.CheckBoxes = True
        
        'INICIO DEL DOCUMENTO DE CONTROL
'        Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
'
'        Dim crParamDef  As CRAXDRT.ParameterFieldDefinition
'
'        Dim objCrystal  As New CRAXDRT.Application
'
'        Dim RutaReporte As String

        RutaReporte = "C:\Admin\Nordi\DoctoControl.rpt"

        Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
        Set crParamDefs = VReporte.ParameterFields

        For Each crParamDef In crParamDefs

            Select Case crParamDef.ParameterFieldName

                Case "mesa"
                    crParamDef.AddCurrentValue str(vPlato)
            End Select

        Next

        LimpiaParametros oCmdEjec
        oCmdEjec.CommandType = adCmdStoredProc
        oCmdEjec.CommandText = "SPPRINTDESPACHO"
        'oCmdEjec.CommandText = "SpPrintComanda"
 
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumSer", adChar, adParamInput, 3, Me.lblSerie.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumFac", adDouble, adParamInput, , Me.lblNumero.Caption)

        Set rsd = oCmdEjec.Execute
    
        'SUB REPORTE
        Dim VReporteS As New CRAXDRT.Report

        VReporte.DataBase.SetDataSource rsd, , 1

        Set VReporteS = VReporte.OpenSubreport("PAGOS")
    
        Set crParamDefs = VReporteS.ParameterFields

        For Each crParamDef In crParamDefs

            Select Case crParamDef.ParameterFieldName

                Case "Pm-ado.NROCOMANDA"
                    crParamDef.AddCurrentValue Me.lblSerie.Caption + "-" + Me.lblNumero.Caption
            End Select

        Next
    
        VReporte.OpenSubreport("PAGOS").DataBase.LogOnServer "p2sodbc.dll", "DSN_DATOS", "bdatos", "sa", claveAcceso
    
        VReporte.PrintOut False, 1, , 1, 1
        Set objCrystal = Nothing
        Set VReporte = Nothing

        Exit Sub

printe:
        MostrarErrores Err
    End If

End Sub

Private Sub cmdEnviarDelivery_Click()
    'AQUI CARGA LA TARIFA DE ENVIO DE DELIVERY
    oCmdEjec.CommandText = "SP_DELIVERY_CARGATARIFAPEDIDO"
    LimpiaParametros oCmdEjec

    Dim orsT As ADODB.Recordset

    Set orsT = oCmdEjec.Execute(, Array(LK_CODCIA, frmDeliveryApp.lblSerie.Caption, frmDeliveryApp.lblNumero.Caption))
            
    If Not orsT.EOF Then
        cMontoTarifa = orsT!tarifa
        cMontoDescuento = orsT!descuento

    End If
    
    frmDeliveryEnviar.gTOTAL = Me.lblTot.Caption
    frmDeliveryEnviar.gMOVILIDAD = cMontoTarifa
    frmDeliveryEnviar.gDESCUENTO = cMontoDescuento

    If Me.DatDireccion.BoundText = "" Then
        frmDeliveryEnviar.gIDDIR = -1
    Else
        frmDeliveryEnviar.gIDDIR = Me.DatDireccion.BoundText

    End If

    frmDeliveryEnviar.Show vbModal

    If frmDeliveryEnviar.vGraba Then Unload Me

End Sub

Private Sub cmdFamAnt_Click()
Dim ini, fin, f As Integer
If vPagActFam = 2 Then
    ini = 1
    fin = ini * 14
ElseIf vPagActFam = 1 Then
    Exit Sub
Else
    FF = vPagActFam - 1
    ini = (14 * FF) - 13
    fin = 14 * FF
End If

For f = ini To fin
    Me.cmdFam(f).Visible = True
Next
If vPagActFam > 1 Then
vPagActFam = vPagActFam - 1
    If vPagActFam = 1 Then: Me.cmdFamAnt.Enabled = False
    
    Me.cmdFamSig.Enabled = True
End If
End Sub

Private Sub cmdFamSig_Click()
Dim ini, fin, f As Integer
If vPagActFam = 1 Then
    ini = 1
    fin = ini * 14
ElseIf vPagActFam = vPagTotFam Then
    Exit Sub
Else
    ini = (14 * vPagActFam) - 13
    fin = 14 * vPagActFam
End If

For f = ini To fin
    Me.cmdFam(f).Visible = False
Next
If vPagActFam < vPagTotFam Then
vPagActFam = vPagActFam + 1
    If vPagActFam = vPagTotFam Then: Me.cmdFamSig.Enabled = False
    
    Me.cmdFamAnt.Enabled = True
End If
End Sub

Private Sub cmdFormaPago_Click()
   
    'AQUI CARGA LA TARIFA DE ENVIO DE DELIVERY
''''    oCmdEjec.CommandText = "SP_DELIVERY_CARGATARIFAPEDIDO"
''''    LimpiaParametros oCmdEjec
''''
''''    Dim orsT As ADODB.Recordset
''''
''''    Set orsT = oCmdEjec.Execute(, Array(LK_CODCIA, frmDeliveryApp.lblSerie.Caption, frmDeliveryApp.lblNumero.Caption))
''''
''''    If Not orsT.EOF Then
''''        cMontoTarifa = orsT!tarifa
''''    End If
''''
''''    frmFacComandaFP.gDELIVERY = True
''''    frmFacComandaFP.Show vbModal
    
frmFacComandaFP2.gDELIVERY = True
'frmFacComandaFP2.lblTotalPagar.Caption = FormatCurrency(Me.lblTot.Caption, 2)
frmFacComandaFP2.Show vbModal

End Sub

Private Sub cmdLimpiar_Click()
Me.lblTexto.Caption = ""
End Sub

Private Sub cmdNum_Click(index As Integer)
Me.lblTexto.Caption = Me.lblTexto.Caption & Me.cmdnum(index).Caption
End Sub

Private Sub cmdPagar_Click()

    If MsgBox("¿Desea continuar con la operación?.", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub
    LimpiaParametros oCmdEjec

    Dim oMSN As String

    oMSN = ""
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.CommandText = "SP_DELIVERY_PAGAR"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIECOMANDA", adVarChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMEROCOMANDA", adBigInt, adParamInput, , Me.lblNumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@exito", adVarChar, adParamOutput, 300, oMSN)
    oCmdEjec.Execute

    oMSN = oCmdEjec.Parameters("@exito").Value

    If Len(Trim(oMSN)) <> 0 Then
        MsgBox oMSN, vbCritical, Pub_Titulo
    Else
        MsgBox "Datos Almacenados Correctamente.", vbInformation, Pub_Titulo
        Unload Me
    End If

End Sub

Private Sub cmdPlato_Click(index As Integer)

    If Len(Trim(Me.lblCliente.Caption)) = 0 Then
        MsgBox "Debe ingresar el Cliente.", vbInformation, Pub_Titulo

        Exit Sub

    End If

    Dim oRStemp As ADODB.Recordset

    'Varificando insumos del plato
    Dim msn     As String

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpDevuelveInsumosxPlato"
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodPlato", adDouble, adParamInput, , CDbl(Me.cmdPlato(index).Tag))
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@mensaje", adVarChar, adParamOutput, 300, msn)

    Dim vstrmin  As String 'variable para capturar los insumos

    Dim vstrcero As String 'variabla para capturar insumos en cero

    Dim vmin     As Boolean 'minimo

    Dim vcero    As Boolean 'stock minimo

    vmin = False
    vcero = False

    Set oRStemp = oCmdEjec.Execute

    If Not oRStemp.EOF Then

        Do While Not oRStemp.EOF

            If oRStemp!sa <= 0 Or (oRStemp!sa - oRStemp!ei) < 0 Then
                vcero = True

                'MsgBox "Algunos insumos del Plato no estan disponibles", vbCritical, NombreProyecto
                If Len(vstrcero) = 0 Then
                    vstrcero = Trim(oRStemp!nm)
                Else
                    vstrcero = vstrcero & vbCrLf & Trim(oRStemp!nm)
                End If

                'Exit Sub
        
            ElseIf (oRStemp!sa - oRStemp!ei) <= oRStemp!sm Then
                vmin = True

                'MsgBox "Algunos insumos del Plato estan el el Minimó permitido", vbInformation, NombreProyecto
                If Len(vstrmin) = 0 Then
                    vstrmin = Trim(oRStemp!nm)
                Else
                    vstrmin = vstrmin & vbCrLf & Trim(oRStemp!nm)
                End If

                'Exit Do
            End If

            'c = c + 1
            oRStemp.MoveNext
        Loop

        'Else
        '    MsgBox "El plato no tiene insumos", vbCritical, NombreProyecto
    End If

    '====ACA SE VE SI CONTROLA INSUMOS==========
    If LK_CONTROLA_STOCK = "A" Then
        If vmin Then
            MsgBox "Los siguientes insumos del Plato estan el el Minimo permitido" & vbCrLf & vstrmin, vbInformation, NombreProyecto
        End If
    
        If vcero Then
            MsgBox "Algunos insumos del Plato no estan disponibles" & vbCrLf & vstrcero, vbCritical, NombreProyecto
        End If
    
        If vcero Then Exit Sub
    Else
    End If

    '====ACA SE VE SI CONTROLA INSUMOS==========

    If VNuevo Then
        If Me.lvPlatos.ListItems.count = 0 Then
            'obteniendo precio
            oRsPlatos.Filter = "Codigo = '" & Me.cmdPlato(index).Tag & "'"
       
            If AgregaPlato(Me.cmdPlato(index).Tag, 1, FormatNumber(oRsPlatos!PRECIO, 2), oRsPlatos!PRECIO, "", "", 0, Me.lblCliente.Caption, 0) Then
        
                With Me.lvPlatos.ListItems.Add(, , Me.cmdPlato(index).Caption, Me.ilPedido.ListImages.Item(1).key, Me.ilPedido.ListImages.Item(1).key)
                    .Tag = Me.cmdPlato(index).Tag
                    .Checked = True
                    .SubItems(2) = " "
                    .SubItems(3) = FormatNumber(1, 2)
                    .SubItems(4) = FormatNumber(oRsPlatos!PRECIO, 2)
                    .SubItems(5) = FormatNumber(val(.SubItems(3)) * val(.SubItems(4)), 2)
                    .SubItems(6) = 0
                    .SubItems(7) = 0   'linea nueva
                    .SubItems(8) = vMaxFac
                    .SubItems(9) = 0
                    VNuevo = False
       
                    oRsPlatos.Filter = ""
                    oRsPlatos.MoveFirst

                End With

                vEstado = "O"
                Me.cmdClienteEdit.Visible = True
                Me.FraCliente.Enabled = False
                
                adicionarBolsa Me.cmdPlato(index).Tag, Me.lblSerie.Caption, Me.lblNumero.Caption, Me.lblCliente.Caption
                
                CargarComanda LK_CODCIA, Me.lblSerie.Caption, Me.lblNumero.Caption
                'Me.DatZona.BoundText = Me.DatDireccion.BoundText
                Me.lvPlatos.ListItems(Me.lvPlatos.ListItems.count).Selected = True
            End If
        End If

    Else

        Dim DD As Integer

        oRsPlatos.Filter = "Codigo = '" & Me.cmdPlato(index).Tag & "'"

        If AgregaPlato(Me.cmdPlato(index).Tag, 1, FormatNumber(oRsPlatos!PRECIO, 2), oRsPlatos!PRECIO, "", Me.lblSerie.Caption, Me.lblNumero.Caption, Me.lblCliente.Caption, 0, DD) Then
    
            With Me.lvPlatos.ListItems.Add(, , Me.cmdPlato(index).Caption, Me.ilPedido.ListImages.Item(1).key, Me.ilPedido.ListImages.Item(1).key)
                .Tag = Me.cmdPlato(index).Tag
                .Checked = True
                .SubItems(3) = FormatNumber(1, 2)
                'obteniendo precio
                oRsPlatos.Filter = "Codigo = '" & Me.cmdPlato(index).Tag & "'"

                If Not oRsPlatos.EOF Then: .SubItems(4) = FormatNumber(oRsPlatos!PRECIO, 2)
                .SubItems(5) = FormatNumber(val(.SubItems(3)) * val(.SubItems(4)), 2)
                .SubItems(6) = DD
                .SubItems(7) = 0   'linea nueva
                .SubItems(8) = vMaxFac
                .SubItems(9) = 0
            End With

            oRsPlatos.Filter = ""
            oRsPlatos.MoveFirst
            
            adicionarBolsa Me.cmdPlato(index).Tag, Me.lblSerie.Caption, Me.lblNumero.Caption, Me.lblCliente.Caption
            
            CargarComanda LK_CODCIA, Me.lblSerie.Caption, Me.lblNumero.Caption
            'Me.DatZona.BoundText = Me.DatDireccion.BoundText
            Me.lvPlatos.ListItems(Me.lvPlatos.ListItems.count).Selected = True
        End If
    End If

    If Me.lvPlatos.ListItems.count <> 0 Then
        'Me.lvPlatos.SelectedItem = Nothing
        Me.lblTot.Caption = FormatCurrency(sumatoria, 2)
        Me.lblItems.Caption = "Items: " & Me.lvPlatos.ListItems.count
        Me.lvPlatos.ListItems(Me.lvPlatos.ListItems.count).Selected = True
        
           
    oCmdEjec.CommandText = "USP_COMANDA_ACTUALIZA_PAGOSDELIVERY"
    LimpiaParametros oCmdEjec
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adVarChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adDouble, adParamInput, , Me.lblNumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TOTAL", adDouble, adParamInput, , Me.lblTot.Caption)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adDouble, adParamInput, , Me.DatZona.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adDouble, adParamInput, , Me.DatZona.BoundText)
    
    oCmdEjec.Execute
    
    
    End If
  
    For c = 1 To Me.lvPlatos.ListItems.count
        'If Me.lvPlatos.ListItems(c).Checked Then
        Me.lvPlatos.ListItems(c).Selected = False
        'End If
    Next
    
End Sub

Private Sub cmdPlatoAnt_Click()
Dim ini, fin, f As Integer
If vPagActFam = 2 Then
    ini = 1
    fin = 22
ElseIf vPagActPla = 1 Then
    Exit Sub
Else
    FF = vPagActPla - 1
    ini = (22 * FF) - 21
    fin = 22 * FF
End If

For f = ini To fin
    Me.cmdPlato(f).Visible = True
Next
If vPagActPla > 1 Then
vPagActPla = vPagActPla - 1
    If vPagActPla = 1 Then: Me.cmdPlatoAnt.Enabled = False
    
    Me.cmdPlatoSig.Enabled = True
End If
End Sub

Private Sub cmdPlatoSig_Click()
Dim ini, fin, f As Integer
If vPagActPla = 1 Then
    ini = 1
    fin = 22
ElseIf vPagActPla = vPagTotPla Then
    Exit Sub
Else
    ini = (22 * vPagActPla) - 21
    fin = 22 * vPagActPla
End If

For f = ini To fin
    Me.cmdPlato(f).Visible = False
Next
If vPagActPla < vPagTotPla Then
vPagActPla = vPagActPla + 1
    If vPagActPla = vPagTotPla Then: Me.cmdPlatoSig.Enabled = False
    
    Me.cmdPlatoAnt.Enabled = True
End If
End Sub

Private Sub cmdPorcion_Click()

    Dim i, c As Integer

    Dim xPRODselecccionados As String

    xPRODselecccionados = ""

    Dim xPROD1, xPRE1 As Double

    Dim xPROD2, xPRE2 As Double

    Dim xSEC1, xSEC2 As Integer

    xPROD1 = 0
    xPROD2 = 0

    c = 0

    For i = 1 To Me.lvPlatos.ListItems.count

        If Me.lvPlatos.ListItems(i).Selected Then
            c = c + 1
        End If

    Next

    If c = 2 Then

        For i = 1 To Me.lvPlatos.ListItems.count

            If Me.lvPlatos.ListItems(i).Selected Then
                xPRODselecccionados = xPRODselecccionados & Me.lvPlatos.ListItems(i).Tag & ","
            End If

        Next
        
        Dim orsSeleccionados As ADODB.Recordset

        oCmdEjec.CommandText = "SP_VERIFICA_PRODUCTO_MEDIAPORCION"
        LimpiaParametros oCmdEjec
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PRODUCTOS", adVarChar, adParamInput, 100, xPRODselecccionados)
    
        Set orsSeleccionados = oCmdEjec.Execute
        
        If Len(Trim(orsSeleccionados!Prod)) <> 0 Then
            MsgBox "Los siguientes articulos no estan habilitados para media porcion." + vbCrLf + orsSeleccionados!Prod, vbInformation, Pub_Titulo
        Else

            'AQUI PROCESA LAS MEDIAS PORCIONES DE LOS PRODUCTOS SELECCIONADOS
            For i = 1 To Me.lvPlatos.ListItems.count

                If Me.lvPlatos.ListItems(i).Selected Then
                    If xPROD1 = 0 Then
                        xPROD1 = Me.lvPlatos.ListItems(i).Tag
                        xPRE1 = Me.lvPlatos.ListItems(i).SubItems(4)
                        xSEC1 = Me.lvPlatos.ListItems(i).SubItems(6)
                    Else
                        xPROD2 = Me.lvPlatos.ListItems(i).Tag
                        xPRE2 = Me.lvPlatos.ListItems(i).SubItems(4)
                        xSEC2 = Me.lvPlatos.ListItems(i).SubItems(6)
                    End If
                
                End If

            Next
                
            'SEGUNDA VALIDACION
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SP_VERIFICA_PRODUCTO_MEDIAPORCION2"
        
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adChar, adParamInput, 3, Me.lblSerie.Caption)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , Me.lblNumero.Caption)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSEC", adInteger, adParamInput, , xSEC1)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , xPROD1)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSECm", adInteger, adParamInput, , xSEC2)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTOm", adBigInt, adParamInput, , xPROD2)
            
            Dim ORSv2 As ADODB.Recordset

            Set ORSv2 = oCmdEjec.Execute
            
            If Len(Trim(ORSv2!Prod)) <> 0 Then
                MsgBox "DEMO" & vbCrLf & ORSv2!Prod
            Else

                LimpiaParametros oCmdEjec
                oCmdEjec.CommandText = "SP_PEDIDO_PROCESAR_MEDIAPORCION"
        
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adChar, adParamInput, 3, Me.lblSerie.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , Me.lblNumero.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)

                If xPRE1 >= xPRE2 Then
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSEC", adInteger, adParamInput, , xSEC1)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , xPROD1)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSECm", adInteger, adParamInput, , xSEC2)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTOm", adBigInt, adParamInput, , xPROD2)
                Else
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSEC", adInteger, adParamInput, , xSEC2)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , xPROD2)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSECm", adInteger, adParamInput, , xSEC1)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTOm", adBigInt, adParamInput, , xPROD1)
                End If

                oCmdEjec.Execute
            End If
            CargarComanda LK_CODCIA, Me.lblSerie.Caption, Me.lblNumero.Caption
            'nuevo 09-08-18  PARA RECALCULAR 1/2 PORCION
            LimpiaParametros oCmdEjec
             oCmdEjec.CommandText = "USP_COMANDA_ACTUALIZA_PAGOSDELIVERY"
    LimpiaParametros oCmdEjec
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adVarChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adDouble, adParamInput, , Me.lblNumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TOTAL", adDouble, adParamInput, , Me.lblTot.Caption)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adDouble, adParamInput, , Me.DatZona.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adDouble, adParamInput, , Me.DatZona.BoundText)
    
    oCmdEjec.Execute
            
        End If

    Else
        MsgBox "Seleccion incorrecta."
    End If

    ''    If Me.lvPlatos.ListItems.count = 0 Then Exit Sub
    ''
    ''    LimpiaParametros oCmdEjec
    ''    oCmdEjec.CommandText = "SP_PEDIDO_VALIDAPORCION"
    ''
    ''    Dim oRSp As ADODB.Recordset
    ''
    ''    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    ''    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , Me.lvPlatos.SelectedItem.Tag)
    ''
    ''    Set oRSp = oCmdEjec.Execute
    ''
    ''    If CBool(oRSp!porc) Then
    ''
    ''        frmPedidoMediaPorcion.gIDproductoOriginal = Me.lvPlatos.SelectedItem.Tag
    ''        frmPedidoMediaPorcion.Show vbModal
    ''
    ''        Dim xPreAct As Double, xSecAct As Integer, xCODprodACT As Double
    ''
    ''        xPreAct = Me.lvPlatos.SelectedItem.SubItems(4)
    ''        xSecAct = Me.lvPlatos.SelectedItem.SubItems(6)
    ''        xCODprodACT = Me.lvPlatos.SelectedItem.Tag
    ''
    ''        If frmPedidoMediaPorcion.gACEPTAR Then
    ''            AgregarDesdeBuscador1 frmPedidoMediaPorcion.gIDproducto, frmPedidoMediaPorcion.gPRODUCTO, IIf(xPreAct > frmPedidoMediaPorcion.gPRECIO, xPreAct, frmPedidoMediaPorcion.gPRECIO), 0.5, xCODprodACT, xSecAct
    ''        End If
        
    ''        On Error GoTo Porcion
    ''
    ''        LimpiaParametros oCmdEjec
    ''        oCmdEjec.CommandText = "SP_PEDIDO_CONVERTIR_PORCION"
    ''
    ''        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    ''        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    ''        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, 2, Me.lblNumero.Caption)
    ''        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adVarChar, adParamInput, 3, Me.lblSerie.Caption)
    ''        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSEC", adInteger, adParamInput, , Me.lvPlatos.SelectedItem.SubItems(6))
    ''        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , Me.lvPlatos.SelectedItem.Tag)
    ''        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PRECIO", adDouble, adParamInput, , oRSp!PPORC)
    ''        oCmdEjec.Execute
    ''
    ''
    ''        Me.lvPlatos.SelectedItem.SubItems(4) = Format(oRSp!PPORC, "#####.#0")
    ''        Me.lvPlatos.SelectedItem.SubItems(3) = 0.5
    ''        Me.lvPlatos.SelectedItem.SubItems(5) = Format(oRSp!PPORC * 0.5, "#####.#0")
    ''
    ''        Me.lblTot.Caption = FormatCurrency(sumatoria, 2)
    ''
    ''        Exit Sub
    ''Porcion:
    ''        MsgBox Err.Description, vbCritical, Pub_Titulo
       
    ''    Else
    ''        MsgBox "El Producto no permite porcion.", vbCritical, Pub_Titulo
    ''    End If

End Sub

Private Sub cmdPrecio_Click()

    If MsgBox("no se puede modificar el Precio" & vbCrLf & "¿Desea ingresar la clave?", vbQuestion + vbYesNo, NombreProyecto) = vbYes Then
        frmClaveCaja.Show vbModal

        If frmClaveCaja.vAceptar Then

            Dim vS As String

            If VerificaPass(frmClaveCaja.vUSUARIO, frmClaveCaja.vClave, vS) Then
                CambiaPrecio
            Else
                MsgBox "Clave incorrecta", vbCritical, NombreProyecto
            End If
        End If
    End If

End Sub

Private Sub cmdPunto_Click()
If Len(Me.lblTexto.Caption) <> 0 Then
    If InStr(Trim(Me.lblTexto.Caption), ".") > 1 Then
        Exit Sub
    Else
        Me.lblTexto.Caption = Me.lblTexto.Caption & Me.cmdPunto.Caption
    End If
End If
End Sub

Private Sub cmdRepartidor_Click()
  If Me.lvPlatos.ListItems.count = 0 Then
      '  MsgBox "Debe agregar un articulo para actualizar observaciones", vbInformation, Pub_Titulo
        Exit Sub
    End If
    
    frmDeliveryCambiaRepartidor.gSerie = Me.lblSerie.Caption
    frmDeliveryCambiaRepartidor.gNumero = Me.lblNumero.Caption
    frmDeliveryCambiaRepartidor.Show vbModal
End Sub

Private Sub cmdSubFam_Click(index As Integer)
Me.cmdPlatoAnt.Enabled = False
Me.cmdPlatoSig.Enabled = False
vColor = index
Me.cmdPlatoAnt.Enabled = False
oRsPlatos.Filter = "CodFam='" & vCodFam & "' and CodSubFam = '" & Me.cmdSubFam(index).Tag & "'"
   For i = 1 To Me.cmdPlato.count - 1
        Unload Me.cmdPlato(i)
    Next

FiltrarPlatos oRsPlatos.RecordCount, oRsPlatos
oRsSubFam.Filter = ""
oRsSubFam.MoveFirst
End Sub

Private Sub cmdFam_Click(index As Integer)
    Me.cmdSubFamAnt.Enabled = False
    Me.cmdSubFamSig.Enabled = False
    vValorActFam = index
    oRsSubFam.Filter = "CodFam='" & cmdFam(index).Tag & "'"
    vCodFam = Me.cmdFam(index).Tag 'Linea Nueva

    If oRsSubFam.RecordCount <> 0 Then
        FiltarSubFamilias oRsSubFam.RecordCount, oRsSubFam
    End If

    For i = 1 To Me.cmdPlato.count - 1
        Unload Me.cmdPlato(i)
    Next

    Me.cmdPlatoAnt.Enabled = False
    Me.cmdPlatoSig.Enabled = False
    oRsSubFam.Filter = ""
    oRsSubFam.MoveFirst

End Sub

Private Sub cmdSubFamAnt_Click()
Dim ini, fin, f As Integer
If vPagActSubFam = 2 Then
    ini = 1
    fin = ini * 14
ElseIf vPagActSubFam = 1 Then
    Exit Sub
Else
    FF = vPagActSubFam - 1
    ini = (14 * FF) - 13
    fin = 14 * FF
End If
 
For f = ini To fin
    Me.cmdSubFam(f).Visible = True
Next
If vPagActSubFam > 1 Then
vPagActSubFam = vPagActSubFam - 1
    If vPagActSubFam = 1 Then: Me.cmdSubFamAnt.Enabled = False
    Me.cmdSubFamSig.Enabled = True
End If
End Sub

Private Sub cmdSubFamSig_Click()

Dim ini, fin, f As Integer
If vPagActSubFam = 1 Then
    ini = 1
    fin = 14
ElseIf vPagActSubFam = vPagTotSubFam Then
    Exit Sub
Else
    ini = (14 * vPagActSubFam) - 13
    fin = 14 * vPagActSubFam
End If
 
For f = ini To fin
    Me.cmdSubFam(f).Visible = False
Next
If vPagActSubFam < vPagTotSubFam Then
    vPagActSubFam = vPagActSubFam + 1
    If vPagActSubFam = vPagTotSubFam Then: Me.cmdSubFamSig.Enabled = False
    Me.cmdSubFamAnt.Enabled = True
End If
End Sub

Private Sub DatDireccion_Change()
'''    Me.datZona.BoundText = Me.DatDireccion.BoundText
'''
'''    Dim orsT As ADODB.Recordset
'''
'''    oRSdir.Filter = ""
'''
'''    If Me.DatDireccion.BoundText = "" Then Exit Sub
'''    oRSdir.Filter = "ZN2=" & Me.DatDireccion.BoundText & " AND DIR='" & Me.DatDireccion.Text & "'"
'''
'''    If Not oRSdir.EOF Then
'''        ORSurb.Filter = "ZN=" & oRSdir!ZONA
'''
'''        If Not ORSurb.EOF Then
'''            Me.txtUrb.Text = ORSurb!urb
'''            Me.lblUrb.Caption = ORSurb!IDEURB
'''        End If
'''    End If
    'Me.DatZona.BoundText = Me.DatDireccion.BoundText

    Dim orsT As ADODB.Recordset

    oRSdir.Filter = ""

    If Me.DatDireccion.BoundText = "" Then Exit Sub
    'oRSdir.Filter = "ZN2=" & Me.DatDireccion.BoundText & " AND DIR='" & Me.DatDireccion.Text & "'"
    oRSdir.Filter = "IDEDIR=" & Me.DatDireccion.BoundText '& " AND DIR='" & Me.DatDireccion.Text & "'"
    Me.lblReferencia.Caption = oRSdir!ref

 Me.txtUrb.Text = ""
            Me.lblurb.Caption = "-1"
            
            Me.DatZona.BoundText = oRSdir!zn2
    If Not oRSdir.EOF Then
        ORSurb.Filter = "ZN=" & oRSdir!ZONA

        If Not ORSurb.EOF Then
            Me.txtUrb.Text = ORSurb!urb
            Me.lblurb.Caption = ORSurb!IDEURB
        End If
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then Unload Me
    If KeyCode = vbKeyF5 Then
        frmComandaProductoSearch.gDELIVERY = True
        frmComandaProductoSearch.Show vbModal
    End If

    If KeyCode = vbKeyF3 Then
        frmAsigna.gDELIVERY = True
        frmAsigna.Mostrador = False
        frmAsigna.Show vbModal
    End If
    
 

End Sub

Private Sub Form_Load()
    vIniLeft = 30
    vIniTop = 120
    CargarFamilias
    CargarSubFamilias
    CargarPlatos
    ConfiguraLV
    vBuscar = True

    If Not VNuevo Then
        CargarComanda LK_CODCIA, vnumser, vNumFac
        'Me.cmdMesa.Enabled = True
    Else
        Me.cmdClienteEdit.Visible = False
    End If

    Me.lblItems.Caption = "Items :" & Me.lvPlatos.ListItems.count

    If vPrimero Then
        Me.lvPlatos.CheckBoxes = False
    Else
        Me.lvPlatos.CheckBoxes = True
    End If
    
    If vEstado = "O" Then
        Me.FraCliente.Enabled = False
        Me.cmdClienteEdit.Visible = False
    End If
    
    If LK_USU_STOCK = "A" Then
        cmdEnviarDelivery.Enabled = True
    Else
        cmdEnviarDelivery.Enabled = False
    End If
    
    '  If oRSfp Is Nothing Then
    Dim cMontoTarifa As Double
      
    oCmdEjec.CommandText = "SP_DELIVERY_CARGATARIFAPEDIDO"
    LimpiaParametros oCmdEjec

    Dim orsT As ADODB.Recordset

    Set orsT = oCmdEjec.Execute(, Array(LK_CODCIA, frmDeliveryApp.lblSerie.Caption, frmDeliveryApp.lblNumero.Caption))
            
    If Not orsT.EOF Then
        cMontoTarifa = orsT!tarifa
    End If

''''    Set oRSfp = New ADODB.Recordset
''''    oRSfp.CursorType = adOpenDynamic ' setting cursor type
''''    oRSfp.Fields.Append "idformapago", adBigInt
''''    oRSfp.Fields.Append "formapago", adVarChar, 120
''''    oRSfp.Fields.Append "moneda", adVarChar, 1
''''    oRSfp.Fields.Append "monto", adDouble
''''
''''    oRSfp.Fields.Refresh
''''    oRSfp.Open
''''
''''    oRSfp.AddNew
''''    oRSfp!idformapago = 1
''''    oRSfp!formapago = "CONTADO"
''''    oRSfp!moneda = "S"
''''    oRSfp!monto = 0 'CDbl(Me.lblTot.Caption) + cMontoTarifa
''''    oRSfp.Update

 Set oRSfp = New ADODB.Recordset
    oRSfp.CursorType = adOpenDynamic ' setting cursor type
    oRSfp.Fields.Append "idformapago", adBigInt
    oRSfp.Fields.Append "formapago", adVarChar, 120
    oRSfp.Fields.Append "referencia", adVarChar, 100
    oRSfp.Fields.Append "monto", adDouble
    oRSfp.Fields.Append "tipo", adChar, 1
    oRSfp.Fields.Append "pagacon", adDouble
    oRSfp.Fields.Append "vuelto", adDouble
    oRSfp.Fields.Append "diascredito", adInteger
    
    oRSfp.Fields.Refresh
    oRSfp.Open
    
    oRSfp.AddNew
    oRSfp!idformapago = 1
    oRSfp!formaPAGO = "CONTADO"
    oRSfp!referencia = ""
    If Len(Trim(Me.lblTot.Caption)) = 0 Then
        oRSfp!monto = 0
    Else
        oRSfp!monto = Me.lblTot.Caption
    End If
    oRSfp!tipo = "E"
    oRSfp!pagacon = 0
    oRSfp!VUELTO = 0
    oRSfp!diascredito = 0
    oRSfp.Update
    ' End If
    gEstaCargando = True
    vFiltro = Leer_Ini(App.Path & "\config.ini", "OPCION_DELIVERY", "0")
    
    
    Me.optFiltro(vFiltro).Value = True
    gEstaCargando = False
End Sub

Private Sub CargarFamilias()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.CommandText = "SpListarFamilias"
    Set oRsFam = oCmdEjec.Execute(, Array(LK_CODCIA))
    'aqui colocar la consulta hacia la data donde devolvera las familias
    'Dim f As Integer
    vfamilia = oRsFam.RecordCount

    'Dim vPri As Boolean
    'vPri = True
    Dim f, c As Integer

    c = 1

    Dim valor As Double

    valor = vfamilia / 14

    pos = InStr(Trim(str(valor)), ".")
    pos2 = Right(Trim(str(valor)), Len(Trim(str(valor))) - pos)
    ent = Left(Trim(str(valor)), pos - 1)

    If ent = "" Then: ent = 0
    If pos2 > 0 Then: vPagTotFam = ent + 1

    If vfamilia >= 1 Then: vPagActFam = 1
    If vfamilia > 14 Then: Me.cmdFamSig.Enabled = True

    For i = 1 To vfamilia
        Load Me.cmdFam(i)
    
        If c <= 3 Then '1 fila
            vIniLeft = vIniLeft + 970
            Me.cmdFam(i).Left = vIniLeft
            Me.cmdFam(i).Top = vIniTop
        ElseIf c <= 7 Then '2º Fila

            'viniLeft = 30
            If c = 4 Then
                vIniLeft = 30
                vIniTop = vIniTop + cmdFamAnt.Height
                Else: vIniLeft = vIniLeft + 970
            End If

        ElseIf c <= 11 Then '3º Fila

            If c = 8 Then
                vIniTop = vIniTop + Me.cmdFam(4).Height
                vIniLeft = 30
                Else: vIniLeft = vIniLeft + 970
            End If

        Else '4º y ultima fila

            If c = 12 Then
                vIniTop = vIniTop + Me.cmdFam(11).Height
                vIniLeft = 30
                Else: vIniLeft = vIniLeft + 970
            End If
        End If

        Me.cmdFam(i).Left = vIniLeft
        Me.cmdFam(i).Top = vIniTop
        'Me.cmdFam(i).Visible = vPri
        Me.cmdFam(i).Visible = True
        Me.cmdFam(i).Caption = Trim(oRsFam!Familia)
        Me.cmdFam(i).Tag = oRsFam!NUMERO
   
        '    If c <= 14 Then
        '        Me.cmdFam(i).Visible = True
        '    Else
        '        Me.cmdFam(i).Visible = False
        '    End If
        oRsFam.MoveNext

        If c = 14 Then
            '        vPri = False
            c = 1
            'vuelve a empezar
            vIniLeft = 30
            vIniTop = 120
        Else
            c = c + 1
        End If
   
    Next

End Sub



Private Sub optFiltro_Click(index As Integer)
vFiltro = index
If gEstaCargando Then Exit Sub
Me.txtCliente.SetFocus
Me.txtCliente.SelStart = 0
Me.txtCliente.SelLength = Len(Me.txtCliente.Text)
End Sub

Private Sub txtCliente_Change()
  vBuscar = True
  loc_key = -1
End Sub

Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo SALE

    Dim strFindMe As String

    Dim itmFound  As Object ' ListItem    ' Variable FoundItem.

    If KeyCode = 40 Then  ' flecha abajo
        loc_key = loc_key + 1

        If loc_key > Me.lvCliente.ListItems.count Then loc_key = lvCliente.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 38 Then
        loc_key = loc_key - 1

        If loc_key < 1 Then loc_key = 1
        GoTo posicion
    End If

    If KeyCode = 34 Then
        loc_key = loc_key + 17

        If loc_key > lvCliente.ListItems.count Then loc_key = lvCliente.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 33 Then
        loc_key = loc_key - 17

        If loc_key < 1 Then loc_key = 1
        GoTo posicion
    End If

    If KeyCode = 27 Then
        Me.lvCliente.Visible = False
        'Me.txtDescExterno.Text = ""
        '        Me.lblDocumento.Caption = ""
        '        Me.lblTelefonos.Caption = ""
    End If

    GoTo fin
posicion:
    lvCliente.ListItems.Item(loc_key).Selected = True
    lvCliente.ListItems.Item(loc_key).EnsureVisible
    
    'Me.txtDescExterno.SelStart = Len(Me.txtDescExterno.Text)
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then
        If vBuscar Then
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SP_DELIVERY_CLIENTE_SEARCH"
            oCmdEjec.CommandType = adCmdStoredProc

            Dim orsC As ADODB.Recordset

            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPO", adInteger, adParamInput, , vFiltro)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SEARCH", adVarChar, adParamInput, 80, Trim(Me.txtCliente.Text))
    
            Set orsC = oCmdEjec.Execute
            Me.lvCliente.ListItems.Clear

            Dim oITEM As Object

            If orsC.RecordCount <> 0 Then
                Me.lvCliente.Visible = True
                loc_key = 1
                vBuscar = False

                Do While Not orsC.EOF
                    Set oITEM = Me.lvCliente.ListItems.Add(, , orsC!cliente)
                    oITEM.Tag = orsC!IDE
                    oITEM.SubItems(1) = Trim(orsC!DNI) 'caleta
                    orsC.MoveNext
                Loop

                Me.lvCliente.ListItems(loc_key).Selected = True
            Else
                Me.lvCliente.Visible = False
                loc_key = 0

                If MsgBox("Cliente No existe. ¿Desea crearlo?.", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub
                frmDeliveryCliente.vTIPO = vFiltro

                If vFiltro = 1 Then
                    If IsNumeric(Me.txtCliente.Text) Then
                        frmDeliveryCliente.txtDni.Text = Me.txtCliente.Text
                    Else
                        frmDeliveryCliente.txtRS.Text = Me.txtCliente.Text
                    End If
                    
                Else
                    frmDeliveryCliente.txtFono.Text = Me.txtCliente.Text
                End If

                frmDeliveryCliente.Show vbModal

                Dim ORSz As ADODB.Recordset

                If frmDeliveryCliente.gIDz <> -1 Then
                    Me.lblCliente.Caption = frmDeliveryCliente.gIDz
                    Me.txtCliente.Text = frmDeliveryCliente.gRS
                    Me.lblDNI.Caption = frmDeliveryCliente.gDNI
                    LimpiaParametros oCmdEjec
                    oCmdEjec.CommandText = "SP_DELIVERY_CLIENTE_DIRECCIONES"
                    
                    Set oRSdir = oCmdEjec.Execute(, Array(LK_CODCIA, Me.lblCliente.Caption))
                    Set Me.DatDireccion.RowSource = Nothing
                    Me.DatDireccion.BoundText = ""
                    Set Me.DatDireccion.RowSource = oRSdir
                    Me.DatDireccion.BoundColumn = oRSdir.Fields(4).Name
                    Me.DatDireccion.ListField = oRSdir.Fields(0).Name
                    Me.DatDireccion.SetFocus
                    Me.DatDireccion.BoundText = frmDeliveryCliente.gDIRECCION
                    
                    'Dim ORSurb As ADODB.Recordset
                    Set ORSurb = oRSdir.NextRecordset
                    Me.txtUrb.Text = ORSurb!urb
                    Me.lblurb.Caption = ORSurb!IDEURB
                    vBuscar = False
                    
                    Set ORSz = oRSdir.NextRecordset
                    Set Me.DatZona.RowSource = ORSz
                    Me.DatZona.BoundColumn = ORSz.Fields(0).Name
                    Me.DatZona.ListField = ORSz.Fields(1).Name
                    Me.DatZona.BoundText = -1
                End If
            End If

        Else
            'BUSCA LAS DIRECCIONES
            'NUEVO SP
            LimpiaParametros oCmdEjec
            'oCmdEjec.Prepared = True
            oCmdEjec.CommandText = "[dbo].[SP_PEDIDOS_VALIDACTUALES]"
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , Me.lvCliente.SelectedItem.Tag)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
            Dim ORSpp As ADODB.Recordset
            Set ORSpp = oCmdEjec.Execute '(, Array(LK_CODCIA, Me.lvCliente.SelectedItem.Tag, LK_FECHA_DIA))
            If Not ORSpp.EOF Then
                If CBool(ORSpp!pendiente) Then
                    If MsgBox("EL CLIENTE TIENE PEDIDOS ANTERIORES. ¿DESEA CONTINUAR?", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then
                    Me.lblCliente.Caption = ""
                    Me.lblDNI.Caption = ""
                    Me.txtCliente.Text = ""
                    Me.lvCliente.ListItems.Clear
                    Me.lvCliente.Visible = False
                    Exit Sub
                    End If
                End If
            End If
            
            Me.lblCliente.Caption = Me.lvCliente.SelectedItem.Tag
            Me.txtCliente.Text = Me.lvCliente.SelectedItem.Text
            Me.lblDNI.Caption = Me.lvCliente.SelectedItem.SubItems(1)
            Me.lvCliente.Visible = False
            Me.lvCliente.ListItems.Clear
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SP_DELIVERY_CLIENTE_DIRECCIONES"

            Set oRSdir = Nothing
            Set oRSdir = oCmdEjec.Execute(, Array(LK_CODCIA, Me.lblCliente.Caption))
            Set Me.DatDireccion.RowSource = Nothing
            Me.DatDireccion.BoundText = ""
            Set Me.DatDireccion.RowSource = oRSdir
            Me.DatDireccion.BoundColumn = oRSdir.Fields(4).Name
            Me.DatDireccion.ListField = oRSdir.Fields(0).Name
            Me.DatDireccion.SetFocus
            
            'Dim ORSurb As ADODB.Recordset
            Set ORSurb = oRSdir.NextRecordset
            '            Me.txtUrb.Text = ORSurb!urb
            '            Me.lblUrb.Caption = ORSurb!IDEURB

            Set ORSz = oRSdir.NextRecordset
            Set Me.DatZona.RowSource = ORSz
            Me.DatZona.BoundColumn = ORSz.Fields(0).Name
            Me.DatZona.ListField = ORSz.Fields(1).Name
            Me.DatZona.BoundText = -1
            
        End If
       
    End If

    ''    KeyAscii = Mayusculas(KeyAscii)
    ''
    ''    If KeyAscii = vbKeyReturn Then
    ''        If vBuscar Then
    ''            LimpiaParametros oCmdEjec
    ''            oCmdEjec.CommandText = "SP_DELIVERY_CLIENTE_SEARCH"
    ''            oCmdEjec.CommandType = adCmdStoredProc
    ''
    ''            Dim orsC As ADODB.Recordset
    ''
    ''            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    ''            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPO", adInteger, adParamInput, , vFiltro)
    ''            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SEARCH", adVarChar, adParamInput, 80, Trim(Me.txtCliente.Text))
    ''
    ''            Set orsC = oCmdEjec.Execute
    ''            Me.lvCliente.ListItems.Clear
    ''
    ''            Dim oITEM As Object
    ''
    ''            If orsC.RecordCount <> 0 Then
    ''                Me.lvCliente.Visible = True
    ''                loc_key = 1
    ''                vBuscar = False
    ''
    ''                Do While Not orsC.EOF
    ''                    Set oITEM = Me.lvCliente.ListItems.Add(, , orsC!cliente)
    ''                    oITEM.Tag = orsC!IDE
    ''                    orsC.MoveNext
    ''                Loop
    ''
    ''                Me.lvCliente.ListItems(loc_key).Selected = True
    ''            Else
    ''                Me.lvCliente.Visible = False
    ''                loc_key = 0
    ''
    ''                If MsgBox("Cliente No existe. ¿Desea crearlo?.", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub
    ''                frmDeliveryCliente.vTIPO = vFiltro
    ''
    ''                If vFiltro = 1 Then
    ''                    frmDeliveryCliente.txtRS.Text = Me.txtCliente.Text
    ''                Else
    ''                    frmDeliveryCliente.txtFono.Text = Me.txtCliente.Text
    ''                End If
    ''
    ''                frmDeliveryCliente.Show vbModal
    ''
    ''                If frmDeliveryCliente.gIDz <> -1 Then
    ''                    Me.lblCliente.Caption = frmDeliveryCliente.gIDz
    ''                    Me.txtCliente.Text = frmDeliveryCliente.gRS
    ''
    ''                    LimpiaParametros oCmdEjec
    ''                    oCmdEjec.CommandText = "SP_DELIVERY_CLIENTE_DIRECCIONES"
    ''
    ''                    Set oRSdir = oCmdEjec.Execute(, Array(LK_CODCIA, Me.lblCliente.Caption))
    ''                    Set Me.DatDireccion.RowSource = Nothing
    ''                    Me.DatDireccion.BoundText = ""
    ''                    Set Me.DatDireccion.RowSource = oRSdir
    ''                    Me.DatDireccion.BoundColumn = oRSdir.Fields(2).Name
    ''                    Me.DatDireccion.ListField = oRSdir.Fields(0).Name
    ''                    Me.DatDireccion.SetFocus
    ''                    Me.DatDireccion.BoundText = frmDeliveryCliente.gDIRECCION
    ''
    ''                    'Dim ORSurb As ADODB.Recordset
    ''                    Set ORSurb = oRSdir.NextRecordset
    ''                    Me.txtUrb.Text = ORSurb!urb
    ''                    Me.lblUrb.Caption = ORSurb!IDEURB
    ''                    vBuscar = False
    ''                End If
    ''            End If
    ''
    ''        Else
    ''            'BUSCA LAS DIRECCIONES
    ''            Me.lblCliente.Caption = Me.lvCliente.SelectedItem.Tag
    ''            Me.txtCliente.Text = Me.lvCliente.SelectedItem.Text
    ''            Me.lvCliente.Visible = False
    ''            Me.lvCliente.ListItems.Clear
    ''            LimpiaParametros oCmdEjec
    ''            oCmdEjec.CommandText = "SP_DELIVERY_CLIENTE_DIRECCIONES"
    ''
    ''            Set oRSdir = Nothing
    ''            Set oRSdir = oCmdEjec.Execute(, Array(LK_CODCIA, Me.lblCliente.Caption))
    ''            Set Me.DatDireccion.RowSource = Nothing
    ''            Me.DatDireccion.BoundText = ""
    ''            Set Me.DatDireccion.RowSource = oRSdir
    ''            Me.DatDireccion.BoundColumn = oRSdir.Fields(2).Name
    ''            Me.DatDireccion.ListField = oRSdir.Fields(0).Name
    ''            Me.DatDireccion.SetFocus
    ''
    ''            'Dim ORSurb As ADODB.Recordset
    ''            Set ORSurb = oRSdir.NextRecordset
    ''            Me.txtUrb.Text = ORSurb!urb
    ''            Me.lblUrb.Caption = ORSurb!IDEURB
    ''
    ''            Dim ORSz As ADODB.Recordset
    ''
    ''            Set ORSz = oRSdir.NextRecordset
    ''            Set Me.datZona.RowSource = ORSz
    ''            Me.datZona.BoundColumn = ORSz.Fields(0).Name
    ''            Me.datZona.ListField = ORSz.Fields(1).Name
    ''            Me.datZona.BoundText = -1
    ''
    ''        End If
    ''
    ''    End If

End Sub

Public Sub AgregarDesdeBuscador(xIDproducto As Double, _
                                xProducto As String, _
                                xPrecio As Double)

    If Len(Trim(Me.lblCliente.Caption)) = 0 Then
        MsgBox "Debe ingresar el Cliente.", vbInformation, Pub_Titulo

        Exit Sub

    End If

    Dim oRStemp As ADODB.Recordset

    'Varificando insumos del plato
    Dim msn     As String

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpDevuelveInsumosxPlato"
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodPlato", adDouble, adParamInput, , xIDproducto)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@mensaje", adVarChar, adParamOutput, 300, msn)

    Dim vstrmin  As String 'variable para capturar los insumos

    Dim vstrcero As String 'variabla para capturar insumos en cero

    Dim vmin     As Boolean 'minimo

    Dim vcero    As Boolean 'stock minimo

    vmin = False
    vcero = False

    Set oRStemp = oCmdEjec.Execute

    If Not oRStemp.EOF Then

        Do While Not oRStemp.EOF

            If oRStemp!sa <= 0 Or (oRStemp!sa - oRStemp!ei) < 0 Then
                vcero = True

                'MsgBox "Algunos insumos del Plato no estan disponibles", vbCritical, NombreProyecto
                If Len(vstrcero) = 0 Then
                    vstrcero = Trim(oRStemp!nm)
                Else
                    vstrcero = vstrcero & vbCrLf & Trim(oRStemp!nm)
                End If

                'Exit Sub
        
            ElseIf (oRStemp!sa - oRStemp!ei) <= oRStemp!sm Then
                vmin = True

                'MsgBox "Algunos insumos del Plato estan el el Minimó permitido", vbInformation, NombreProyecto
                If Len(vstrmin) = 0 Then
                    vstrmin = Trim(oRStemp!nm)
                Else
                    vstrmin = vstrmin & vbCrLf & Trim(oRStemp!nm)
                End If

                'Exit Do
            End If

            'c = c + 1
            oRStemp.MoveNext
        Loop

        'Else
        '    MsgBox "El plato no tiene insumos", vbCritical, NombreProyecto
    End If

    '====ACA SE VE SI CONTROLA INSUMOS==========
    If LK_CONTROLA_STOCK = "A" Then
        If vmin Then
            MsgBox "Los siguientes insumos del Plato estan el el Minimo permitido" & vbCrLf & vstrmin, vbInformation, NombreProyecto
        End If
    
        If vcero Then
            MsgBox "Algunos insumos del Plato no estan disponibles" & vbCrLf & vstrcero, vbCritical, NombreProyecto
        End If
    
        If vcero Then Exit Sub
    Else
    End If

    '====ACA SE VE SI CONTROLA INSUMOS==========

    If VNuevo Then
        If Me.lvPlatos.ListItems.count = 0 Then
            'obteniendo precio
            'oRsPlatos.Filter = "Codigo = '" & Me.cmdPlato(Index).Tag & "'"
       
            If AgregaPlato(xIDproducto, 1, xPrecio, xPrecio, "", "", 0, Me.lblCliente.Caption, 0) Then
        
                With Me.lvPlatos.ListItems.Add(, , xIDproducto, Me.ilPedido.ListImages.Item(1).key, Me.ilPedido.ListImages.Item(1).key)
                    .Tag = Me.cmdPlato(index).Tag
                    .Checked = True
                    .SubItems(2) = " "
                    .SubItems(3) = FormatNumber(1, 2)
                    .SubItems(4) = FormatNumber(xPrecio, 2)
                    .SubItems(5) = FormatNumber(val(.SubItems(3)) * val(.SubItems(4)), 2)
                    .SubItems(6) = 0
                    .SubItems(7) = 0   'linea nueva
                    .SubItems(8) = vMaxFac
                    .SubItems(9) = 0
                    VNuevo = False
       
                    'oRsPlatos.Filter = ""
                    'oRsPlatos.MoveFirst

                End With

                vEstado = "O"
            End If
        End If

    Else

        Dim DD As Integer

        'oRsPlatos.Filter = "Codigo = '" & Me.cmdPlato(Index).Tag & "'"

        If AgregaPlato(xIDproducto, 1, FormatNumber(xPrecio, 2), xPrecio, "", Me.lblSerie.Caption, Me.lblNumero.Caption, Me.lblCliente.Caption, 0, DD) Then
    
            With Me.lvPlatos.ListItems.Add(, , xProducto, Me.ilPedido.ListImages.Item(1).key, Me.ilPedido.ListImages.Item(1).key)
                .Tag = Me.cmdPlato(index).Tag
                .Checked = True
                .SubItems(3) = FormatNumber(1, 2)
                'obteniendo precio
                'oRsPlatos.Filter = "Codigo = '" & Me.cmdPlato(Index).Tag & "'"

                'If Not oRsPlatos.EOF Then:
                .SubItems(4) = FormatNumber(xPrecio, 2)
                .SubItems(5) = FormatNumber(val(.SubItems(3)) * val(.SubItems(4)), 2)
                .SubItems(6) = DD
                .SubItems(7) = 0   'linea nueva
                .SubItems(8) = vMaxFac
                .SubItems(9) = 0
            End With

            ' oRsPlatos.Filter = ""
            'oRsPlatos.MoveFirst
    
        End If
    End If

    If Me.lvPlatos.ListItems.count <> 0 Then
        'Me.lvPlatos.SelectedItem = Nothing
        Me.lblTot.Caption = FormatCurrency(sumatoria, 2)
        Me.lblItems.Caption = "Items: " & Me.lvPlatos.ListItems.count
        Me.lvPlatos.ListItems(Me.lvPlatos.ListItems.count).Selected = True
    End If

    CargarComanda LK_CODCIA, Me.lblSerie.Caption, Me.lblNumero.Caption
  
    For c = 1 To Me.lvPlatos.ListItems.count
        'If Me.lvPlatos.ListItems(c).Checked Then
        Me.lvPlatos.ListItems(c).Selected = False
        'End If
    Next

    Me.lvPlatos.ListItems(Me.lvPlatos.ListItems.count).Selected = True
End Sub

Public Sub AgregarDesdeBuscador1(xIDproducto As Double, _
                                 xProducto As String, _
                                 xPrecio As Double, _
                                 xCANT As Double, xIDproductoACT As Double, _
                                 xNROsecACT As Integer)

    If Len(Trim(Me.lblCliente.Caption)) = 0 Then
        MsgBox "Debe ingresar el Cliente.", vbInformation, Pub_Titulo

        Exit Sub

    End If

    Dim oRStemp As ADODB.Recordset

    'Varificando insumos del plato
    Dim msn     As String

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpDevuelveInsumosxPlato"
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodPlato", adDouble, adParamInput, , xIDproducto)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@mensaje", adVarChar, adParamOutput, 300, msn)

    Dim vstrmin  As String 'variable para capturar los insumos

    Dim vstrcero As String 'variabla para capturar insumos en cero

    Dim vmin     As Boolean 'minimo

    Dim vcero    As Boolean 'stock minimo

    vmin = False
    vcero = False

    Set oRStemp = oCmdEjec.Execute

    If Not oRStemp.EOF Then

        Do While Not oRStemp.EOF

            If oRStemp!sa <= 0 Or (oRStemp!sa - oRStemp!ei) < 0 Then
                vcero = True

                'MsgBox "Algunos insumos del Plato no estan disponibles", vbCritical, NombreProyecto
                If Len(vstrcero) = 0 Then
                    vstrcero = Trim(oRStemp!nm)
                Else
                    vstrcero = vstrcero & vbCrLf & Trim(oRStemp!nm)
                End If

                'Exit Sub
        
            ElseIf (oRStemp!sa - oRStemp!ei) <= oRStemp!sm Then
                vmin = True

                'MsgBox "Algunos insumos del Plato estan el el Minimó permitido", vbInformation, NombreProyecto
                If Len(vstrmin) = 0 Then
                    vstrmin = Trim(oRStemp!nm)
                Else
                    vstrmin = vstrmin & vbCrLf & Trim(oRStemp!nm)
                End If

                'Exit Do
            End If

            'c = c + 1
            oRStemp.MoveNext
        Loop

        'Else
        '    MsgBox "El plato no tiene insumos", vbCritical, NombreProyecto
    End If

    '====ACA SE VE SI CONTROLA INSUMOS==========
    If LK_CONTROLA_STOCK = "A" Then
        If vmin Then
            MsgBox "Los siguientes insumos del Plato estan el el Minimo permitido" & vbCrLf & vstrmin, vbInformation, NombreProyecto
        End If
    
        If vcero Then
            MsgBox "Algunos insumos del Plato no estan disponibles" & vbCrLf & vstrcero, vbCritical, NombreProyecto
        End If
    
        If vcero Then Exit Sub
    Else
    
    End If

    '====ACA SE VE SI CONTROLA INSUMOS==========

    '    If VNuevo Then
    '        If Me.lvPlatos.ListItems.count = 0 Then
    '            'obteniendo precio
    '            'oRsPlatos.Filter = "Codigo = '" & Me.cmdPlato(Index).Tag & "'"
    '
    '            If AgregaPlato(xIDproducto, 1, xPrecio, xPrecio, "", "", 0, Me.lblCliente.Caption, 0) Then
    '
    '                With Me.lvPlatos.ListItems.Add(, , xIDproducto, Me.ilPedido.ListImages.Item(1).Key, Me.ilPedido.ListImages.Item(1).Key)
    '                    .Tag = Me.cmdPlato(Index).Tag
    '                    .Checked = True
    '                    .SubItems(2) = " "
    '                    .SubItems(3) = FormatNumber(1, 2)
    '                    .SubItems(4) = FormatNumber(xPrecio, 2)
    '                    .SubItems(5) = FormatNumber(val(.SubItems(3)) * val(.SubItems(4)), 2)
    '                    .SubItems(6) = 0
    '                    .SubItems(7) = 0   'linea nueva
    '                    .SubItems(8) = vMaxFac
    '                    .SubItems(9) = 0
    '                    VNuevo = False
    '
    '                    'oRsPlatos.Filter = ""
    '                    'oRsPlatos.MoveFirst
    '
    '                End With
    '
    '                vEstado = "O"
    '            End If
    '        End If
    '
    '    Else

    Dim DD As Integer

    'oRsPlatos.Filter = "Codigo = '" & Me.cmdPlato(Index).Tag & "'"

    If AgregaPlato(xIDproducto, xCANT, FormatNumber(xPrecio, 2), xPrecio * xCANT, "", Me.lblSerie.Caption, Me.lblNumero.Caption, Me.lblCliente.Caption, 0, DD) Then
    
    'MODIFICANDO EL PRODUCTO ORIGINAL
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_PEDIDO_ACTUALIZAR_MEDIAPORCION"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adChar, adParamInput, 3, Me.lblSerie.Caption)
    
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , Me.lblNumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSEC", adInteger, adParamInput, , xNROsecACT)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODART", adBigInt, adParamInput, , xIDproductoACT)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CANTIDAD", adDouble, adParamInput, , 0.5)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PRECIO", adDouble, adParamInput, , xPrecio)
    oCmdEjec.Execute
    
        With Me.lvPlatos.ListItems.Add(, , xProducto, Me.ilPedido.ListImages.Item(1).key, Me.ilPedido.ListImages.Item(1).key)
            .Tag = Me.cmdPlato(index).Tag
            .Checked = True
            .SubItems(3) = FormatNumber(xCANT, 2)
            'obteniendo precio
            'oRsPlatos.Filter = "Codigo = '" & Me.cmdPlato(Index).Tag & "'"

            'If Not oRsPlatos.EOF Then:
            .SubItems(4) = FormatNumber(xPrecio, 2)
            .SubItems(5) = FormatNumber(val(.SubItems(3)) * val(.SubItems(4)), 2)
            .SubItems(6) = DD
            .SubItems(7) = 0   'linea nueva
            .SubItems(8) = vMaxFac
            .SubItems(9) = 0
        End With

        ' oRsPlatos.Filter = ""
        'oRsPlatos.MoveFirst
    
    End If

    '    End If

    If Me.lvPlatos.ListItems.count <> 0 Then
        'Me.lvPlatos.SelectedItem = Nothing
        Me.lblTot.Caption = FormatCurrency(sumatoria, 2)
        Me.lblItems.Caption = "Items: " & Me.lvPlatos.ListItems.count
        Me.lvPlatos.ListItems(Me.lvPlatos.ListItems.count).Selected = True
    End If

    CargarComanda LK_CODCIA, Me.lblSerie.Caption, Me.lblNumero.Caption
  
    For c = 1 To Me.lvPlatos.ListItems.count
        'If Me.lvPlatos.ListItems(c).Checked Then
        Me.lvPlatos.ListItems(c).Selected = False
        'End If
    Next

    Me.lvPlatos.ListItems(Me.lvPlatos.ListItems.count).Selected = True
End Sub

Private Sub txtUrb_Change()
vBuscar_u = True
Me.lblurb.Caption = -1
End Sub

Private Sub txtUrb_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo SALE

    Dim strFindMe As String

    Dim itmFound  As Object ' ListItem    ' Variable FoundItem.

    If KeyCode = 40 Then  ' flecha abajo
        loc_key_u = loc_key_u + 1

        If loc_key_u > ListView1.ListItems.count Then loc_key_u = ListView1.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 38 Then
        loc_key_u = loc_key_u - 1

        If loc_key_u < 1 Then loc_key_u = 1
        GoTo posicion
    End If

    If KeyCode = 34 Then
        loc_key_u = loc_key_u + 17

        If loc_key_u > ListView1.ListItems.count Then loc_key_u = ListView1.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 33 Then
        loc_key_u = loc_key_u - 17

        If loc_key_u < 1 Then loc_key_u = 1
        GoTo posicion
    End If

    If KeyCode = 27 Then
        Me.ListView1.Visible = False
        Me.txtUrb.Text = ""
        Me.lblurb.Caption = "-1"
    End If

    GoTo fin
posicion:
    ListView1.ListItems.Item(loc_key_u).Selected = True
    ListView1.ListItems.Item(loc_key_u).EnsureVisible
    'txtRS.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
    txtUrb.SelStart = Len(Me.txtUrb.Text)
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtUrb_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then
        If vBuscar_u Then
            Me.ListView1.ListItems.Clear
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SpListarCliUrbanizaciones"
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SEARCH", adVarChar, adParamInput, 80, Me.txtUrb.Text)
                
            Dim ORSurb As ADODB.Recordset
                
            Set ORSurb = oCmdEjec.Execute

            Dim Item As Object
        
            If Not ORSurb.EOF Then

                Do While Not ORSurb.EOF
                    Set Item = Me.ListView1.ListItems.Add(, , ORSurb!nom)
                    Item.Tag = ORSurb!IDE
                    ORSurb.MoveNext
                Loop

                Me.ListView1.Visible = True
                Me.ListView1.ListItems(1).Selected = True
                loc_key_u = 1
                Me.ListView1.ListItems(1).EnsureVisible
                vBuscar_u = False
            Else

                If MsgBox("Urbanización no encontrada." & vbCrLf & "¿Desea agregarla?", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub
                frmClientesUrbAdd.txtUrb.Text = Me.txtUrb.Text
                frmClientesUrbAdd.Show vbModal

                If frmClientesUrbAdd.gAcepta Then
                    Me.txtUrb.Text = frmClientesUrbAdd.gNombre
                    Me.lblurb.Caption = frmClientesUrbAdd.gIde
                    Me.ListView1.Visible = False
                    vBuscar_u = True
                    'AQUI GRABA
                    LimpiaParametros oCmdEjec
                    oCmdEjec.CommandText = "SP_DELIVERY_ACTUALIZA_URB_CLIENTE"
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , Me.lblCliente.Caption)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDURB", adInteger, adParamInput, , Me.lblurb.Caption)
                    oCmdEjec.Execute
                    'Me.datZona.SetFocus
                End If
            End If
        
        Else
            Me.txtUrb.Text = Me.ListView1.ListItems(loc_key_u).Text
            Me.lblurb.Caption = Me.ListView1.ListItems(loc_key_u).Tag
            
             LimpiaParametros oCmdEjec
                    oCmdEjec.CommandText = "SP_DELIVERY_ACTUALIZA_URB_CLIENTE"
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , Me.lblCliente.Caption)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDURB", adInteger, adParamInput, , Me.lblurb.Caption)
                    oCmdEjec.Execute
                    
            Me.ListView1.Visible = False
            vBuscar = True
            '   Me.datZona.SetFocus
            
            'Me.lvDetalle.SetFocus
        End If
    End If

End Sub

Private Sub adicionarBolsa(cIDarticulo As Integer, cSerie As String, cNumero As Double, cCliente As String)

    Dim oRSbolsa As ADODB.Recordset

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_VALIDA_ARTICULO_BOLSA]"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODART", adBigInt, adParamInput, , cIDarticulo)
                
    Set oRSbolsa = oCmdEjec.Execute
    VNuevo = False

    Dim DD2 As Integer

    If Not oRSbolsa.EOF Then
        If oRSbolsa!AGREGA Then
            If AgregaPlato(oRSbolsa!IDE, oRSbolsa!cant, oRSbolsa!PRE, (oRSbolsa!cant * oRSbolsa!PRE), "", cSerie, cNumero, cCliente, IIf(Len(Trim(Me.lblCliente.Caption)) = 0, 0, Me.lblCliente.Caption), DD2) Then

                With Me.lvPlatos.ListItems.Add(, , oRSbolsa!PRD, Me.ilPedido.ListImages.Item(1).key, Me.ilPedido.ListImages.Item(1).key)
                    .Tag = oRSbolsa!IDE
                    .Checked = True
                    .SubItems(2) = " "
                    .SubItems(3) = FormatNumber(oRSbolsa!cant, 2)
                    .SubItems(4) = FormatNumber(oRSbolsa!PRE, 2)
                    .SubItems(5) = FormatNumber(val(.SubItems(3)) * val(.SubItems(4)), 2)
                    .SubItems(6) = DD2
                    .SubItems(7) = 0   'linea nueva
                    .SubItems(8) = vMaxFac
                    .SubItems(9) = 0
                    VNuevo = False
       

                End With

            End If

        End If

    End If

    'AGREGA BOLSA SEGUN CONFIGURACION DEL ARTICULO - FIN
End Sub

