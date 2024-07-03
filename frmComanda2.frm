VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComanda2 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   ClientHeight    =   7635
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   15675
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmComanda2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   15675
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDescuentos 
      Caption         =   "Descuentos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   8400
      TabIndex        =   47
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdPorcion 
      Caption         =   "1/2 Porcion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   4680
      TabIndex        =   46
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCarac 
      Caption         =   "Carac"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   8400
      TabIndex        =   45
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdCta 
      Caption         =   "Cuenta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8400
      TabIndex        =   40
      Top             =   1320
      Width           =   975
   End
   Begin Crystal.CrystalReport crReporte 
      Left            =   1920
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdBorrar 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12360
      Picture         =   "frmComanda2.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3120
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
      Left            =   11400
      TabIndex        =   15
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
      Left            =   9480
      TabIndex        =   14
      Top             =   3120
      Width           =   1935
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
      Left            =   12360
      TabIndex        =   13
      Top             =   2400
      Width           =   975
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
      Left            =   11400
      TabIndex        =   12
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
      Left            =   10440
      TabIndex        =   11
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
      Left            =   9480
      TabIndex        =   10
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
      Left            =   12360
      TabIndex        =   9
      Top             =   1680
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
      Left            =   11400
      TabIndex        =   8
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
      Left            =   10440
      TabIndex        =   7
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
      Left            =   9480
      TabIndex        =   6
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
      Left            =   12360
      TabIndex        =   1
      Top             =   960
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
      Left            =   11400
      TabIndex        =   5
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
      Left            =   10440
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdNum 
      Appearance      =   0  'Flat
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
      Left            =   9480
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8400
      Picture         =   "frmComanda2.frx":0FB4
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar Plato"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8400
      Picture         =   "frmComanda2.frx":175E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin MSComctlLib.ListView lvPlatos 
      Height          =   3255
      Left            =   90
      TabIndex        =   0
      Top             =   480
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   5741
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      Icons           =   "ilComanda"
      SmallIcons      =   "ilComanda"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cer&rar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   13440
      Picture         =   "frmComanda2.frx":1E49
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2400
      Width           =   2055
   End
   Begin MSComctlLib.ImageList ilComanda 
      Left            =   2880
      Top             =   3840
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
            Picture         =   "frmComanda2.frx":25F3
            Key             =   "Plato"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3105
      Left            =   120
      TabIndex        =   21
      Top             =   4440
      Width           =   3975
      Begin VB.CommandButton cmdFam 
         Caption         =   "Familia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   1320
         TabIndex        =   24
         Top             =   3120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdFamAnt 
         Caption         =   "Anterior"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   30
         Picture         =   "frmComanda2.frx":2CED
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdFamSig 
         Caption         =   "Siguiente"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2940
         Picture         =   "frmComanda2.frx":33D8
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2325
         Width           =   975
      End
   End
   Begin VB.Frame fraSubFam 
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3105
      Left            =   4320
      TabIndex        =   18
      Top             =   4440
      Width           =   4900
      Begin VB.CommandButton cmdSubFam 
         Caption         =   "SubFam"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   3480
         TabIndex        =   25
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdSubFamSig 
         Caption         =   "Siguiente"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3675
         Picture         =   "frmComanda2.frx":3AC3
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2325
         Width           =   1215
      End
      Begin VB.CommandButton cmdSubFamAnt 
         Caption         =   "Anterior"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   30
         Picture         =   "frmComanda2.frx":41AE
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame fraPlatos 
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3100
      Left            =   9360
      TabIndex        =   26
      Top             =   4440
      Width           =   6160
      Begin VB.CommandButton cmdPlatoSig 
         Caption         =   "Siguiente"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4890
         Picture         =   "frmComanda2.frx":4899
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2325
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlato 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Plato"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlatoAnt 
         Caption         =   "Anterior"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   30
         Picture         =   "frmComanda2.frx":4F84
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdPreCuenta 
      Caption         =   "&Pre-Cuenta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   13440
      TabIndex        =   37
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   13440
      Picture         =   "frmComanda2.frx":566F
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton cmdFactura 
      Caption         =   "&Facturar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   13440
      TabIndex        =   38
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label lblComensales 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   14520
      TabIndex        =   44
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lblCliente 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   9360
      TabIndex        =   43
      Top             =   3960
      Width           =   5055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
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
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   5640
      TabIndex        =   42
      Top             =   3915
      Width           =   660
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
      Left            =   6285
      TabIndex        =   41
      Top             =   3840
      Width           =   2070
   End
   Begin VB.Label lblmesa 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   9375
   End
   Begin VB.Label lblItems 
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   120
      TabIndex        =   34
      Top             =   3855
      Width           =   1020
   End
   Begin VB.Label lblNumero 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   14160
      TabIndex        =   32
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblSerie 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   13440
      TabIndex        =   31
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblMozo 
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9480
      TabIndex        =   30
      Tag             =   "1"
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label lblTexto 
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
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   9480
      TabIndex        =   17
      Top             =   480
      Width           =   3855
   End
End
Attribute VB_Name = "frmComanda2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vfamilia As Integer
Public vSubFamilia As Integer
Private vPlato As Integer
Private vIniLeft As Integer
Private vIniTop As Integer
Private vPagActFam, vPagActSubFam, vPagActPla As Integer
Private vPagTotFam, vPagTotSubFam, vPagTotPla As Integer
Private oRsFam As ADODB.Recordset
Private oRsSubFam As ADODB.Recordset
Public oRsPlatos As ADODB.Recordset
Private vValorActFam As Integer
Public VNuevo As Boolean
Public vPrimero As Boolean
 Const vMesa As String = "0"
Public vCodZona As Integer
Public vEstado As String
Private vMesaAnt As String
Public vCodFam As Integer
Private vColor As Integer
Public vMaxFac As Double
 
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

    If VNuevo Then 'nueva comanda

        With oCmdEjec
            .CommandText = "SpRegistrarComanda_MOSTRADOR"
            .Parameters.Append .CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
            .Parameters.Append .CreateParameter("@Usuario", adVarChar, adParamInput, 10, LK_CODUSU)
            .Parameters.Append .CreateParameter("@CodMesa", adVarChar, adParamInput, 10, vMesa)
            .Parameters.Append .CreateParameter("@cp", adInteger, adParamInput, , vcp)
            .Parameters.Append .CreateParameter("@cant", adDouble, adParamInput, , vc)
            .Parameters.Append .CreateParameter("@pre", adDouble, adParamInput, , vpre)
            .Parameters.Append .CreateParameter("@imp", adDouble, adParamInput, , vimp)
            .Parameters.Append .CreateParameter("@d", adVarChar, adParamInput, 50, vd)

            '       .Parameters.Append .CreateParameter("@Total", adCurrency, adParamInput, , Val(Me.lblTot.Tag))
            .Parameters.Append .CreateParameter("@Mozo", adInteger, adParamInput, , CInt(Me.lblMozo.Tag))

            .Parameters.Append .CreateParameter("@NumSer", adChar, adParamOutput, 3, NumSer)
            .Parameters.Append .CreateParameter("@NumFac", adDouble, adParamOutput, , NumFac)
            .Parameters.Append .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
            .Parameters.Append .CreateParameter("@CodFam", adInteger, adParamInput, , vCodFam)

            'PARAMETROS NUEVOS
            
                .Parameters.Append .CreateParameter("@CLIENTE", adVarChar, adParamInput, 120, VcLIENTE) 'Linea nueva
            

            .Parameters.Append .CreateParameter("@COMENSALES", adDouble, adParamInput, , VcOMENSALES)  'Linea nueva
            '.Parameters.Append .CreateParameter("@ZONA", adInteger, adParamInput, , vCodZona)
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
            .CommandText = "SpModificarComanda1"
            .Parameters.Append .CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
            .Parameters.Append .CreateParameter("@Usuario", adVarChar, adParamInput, 10, LK_CODUSU)
            .Parameters.Append .CreateParameter("@CodMesa", adVarChar, adParamInput, 10, vMesa)
            .Parameters.Append .CreateParameter("@cp", adDouble, adParamInput, , vcp) 'julio 11-01-2011
            .Parameters.Append .CreateParameter("@cant", adDouble, adParamInput, , vc)
            .Parameters.Append .CreateParameter("@pre", adDouble, adParamInput, , vpre)
            .Parameters.Append .CreateParameter("@imp", adDouble, adParamInput, , vimp)
            .Parameters.Append .CreateParameter("@d", adVarChar, adParamInput, 50, vd)
        
            '       .Parameters.Append .CreateParameter("@Total", adCurrency, adParamInput, , Val(Me.lblTot.Tag))
            .Parameters.Append .CreateParameter("@Mozo", adInteger, adParamInput, , CInt(Me.lblMozo.Tag))
        
            .Parameters.Append .CreateParameter("@NumSer", adChar, adParamInput, 3, vnumser)
            .Parameters.Append .CreateParameter("@NumFac", adDouble, adParamInput, , vNumFac)
            .Parameters.Append .CreateParameter("@NUMSEC", adInteger, adParamOutput)
            .Parameters.Append .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
            .Parameters.Append .CreateParameter("@CodFam", adInteger, adParamInput, , vCodFam)  'linea nueva
             .Parameters.Append .CreateParameter("@CLIENTE", adVarChar, adParamInput, 120, VcLIENTE) 'Linea nueva
            

            .Parameters.Append .CreateParameter("@COMENSALES", adDouble, adParamInput, , VcOMENSALES)  'Linea nueva
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
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@mesa", adVarChar, adParamInput, 10, vMesa)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@tipo", adBoolean, adParamInput, , 1) '0 cuando es extorno
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumSec", adInteger, adParamInput, , vnumsec)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MaxNumFac", adDouble, adParamOutput, , vMaxFac)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MaxNumOper", adDouble, adParamOutput, , 0)

    oCmdEjec.Execute
    vMaxFac = oCmdEjec.Parameters("@MaxNumFac").Value
    vMaxNumoper = oCmdEjec.Parameters("@MaxNumOper").Value

    'LimpiaParametros oCmdEjec
    'oCmdEjec.CommandText = "SpActualizarPedTrans"
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumSer", adChar, adParamInput, 3, Me.lblSerie.Caption)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumFac", adInteger, adParamInput, , Me.lblNumero.Caption)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumSec", adInteger, adParamInput, , vnumsec)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumOper", adInteger, adParamInput, , vMaxNumoper)
    'oCmdEjec.Execute
    AgregaPlato = True

    Exit Function

ErrorGraba:
    AgregaPlato = False
    MsgBox Err.Description
End Function

Public Sub CargarComanda(vCodCia As String, vCodMesa As String)

    Dim oRsComanda As ADODB.Recordset

    Me.lvPlatos.ListItems.Clear

    Dim vnumser, vMozo As String

    Dim vNumFac, vCodMozo As Integer

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpCargarComanda2"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodMesa", adVarChar, adParamInput, 10, vCodMesa)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 20, LK_CODUSU)

    Set oRsComanda = oCmdEjec.Execute

    If Not oRsComanda.EOF Then
        Me.lblNumero.Caption = oRsComanda.Fields!PED_NUMFAC ' oCmdEjec.Parameters("@NumFac").Value
        Me.lblSerie.Caption = oRsComanda.Fields!PED_numser '  oCmdEjec.Parameters("@NumSer").Value
        Me.lblMozo.Tag = oRsComanda.Fields!PED_CODVEN  'oCmdEjec.Parameters("@CodMozo").Value
        Me.lblMozo.Caption = Trim(oRsComanda.Fields!mozo)  'Trim(oCmdEjec.Parameters("@Mozo").Value)
        Me.lblCliente.Caption = IIf(IsNull(Trim(oRsComanda!cliente)), "", oRsComanda!cliente)
        Me.lblComensales.Caption = Trim(oRsComanda!Comensales)

        '    Me.lblSerie.Tag = oRsComanda!NumFac
        Do While Not oRsComanda.EOF
    
            With Me.lvPlatos.ListItems.Add(, , Trim(oRsComanda!plato), Me.ilComanda.ListImages(1).key, Me.ilComanda.ListImages(1).key)
                .Tag = oRsComanda!CODPLATO
                '.SubItems(1) = iif(oRsComanda!cuenta Trim(oRsComanda!cuenta)
                .SubItems(1) = IIf(IsNull(oRsComanda!cuenta), "", Trim(oRsComanda!cuenta))
                .SubItems(2) = Trim(oRsComanda!DETALLE)
                .SubItems(3) = Format(oRsComanda!Cantidad, "#####.#0")
                .SubItems(4) = Format(oRsComanda!PRECIO, "#####.#0")
                .SubItems(5) = Format(oRsComanda!Importe, "#####.#0")
                .SubItems(6) = oRsComanda!Sec
                .SubItems(7) = oRsComanda!aten
                '.SubItems(7) = oRsComanda!NumFac
                .SubItems(8) = oRsComanda.Fields!PED_NUMFAC
                .SubItems(9) = oRsComanda!aPRO
                .SubItems(10) = oRsComanda!NumFac
                .SubItems(11) = oRsComanda!cantado
                '.SubItems(10) = oRsComanda!PED_NUMFAC
                .SubItems(12) = oRsComanda!Enviar
                .SubItems(13) = oRsComanda!fam

                '.SubItems(10) = oRsComanda!PED_NUMFAC
                If oRsComanda!aPRO = "0" Then .Checked = True
            End With

            'Set itemC = Me.lvPlatos.ListItems.Add(, , Trim(oRsComanda!Plato), Me.ilComanda.ListImages(1).Key, Me.ilComanda.ListImages(1).Key)
            '        itemC.Tag = oRsComanda!CodPlato
            '        itemC.SubItems(1) = Trim(oRsComanda!detalle)
            '        itemC.SubItems(2) = FormatNumber(oRsComanda!Cantidad, 2)
            '        itemC.SubItems(3) = FormatNumber(oRsComanda!Precio, 2)
            '        itemC.SubItems(4) = FormatNumber(oRsComanda!Importe, 2)
            '        itemC.SubItems(5) = oRsComanda!sec
            '        itemC.SubItems(6) = oRsComanda!aten
            '        itemC.SubItems(7) = oRsComanda!NumFac
            '        If oRsComanda!aPRO = "0" Then itemC.Checked = True
            Me.lblTot.Caption = FormatCurrency(sumatoria, 2)
        
            oRsComanda.MoveNext
        Loop

    End If

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

Private Sub elimina(cNumSec As Integer, cCodPlato As Double, cCantidad As Integer)
Dim cCant As Integer
cCant = Me.lvPlatos.ListItems.count

    On Error GoTo eli

    Pub_ConnAdo.BeginTrans
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.CommandText = "SpActualizarPlato"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adVarChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adDouble, adParamInput, , Me.lblNumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numsec", adInteger, adParamInput, , cNumSec)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numsec", adInteger, adParamInput, , Me.lvPlatos.SelectedItem.SubItems(6))
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
    
     If Me.lvPlatos.ListItems.count = 1 Then
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SpLiberarMesa"
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodMesa", adVarChar, adParamInput, 10, vMesa)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodZon", adInteger, adParamInput, , vCodZona)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adVarChar, adParamInput, 3, Me.lblSerie.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , Me.lblNumero.Caption)
        oCmdEjec.Execute

        'Unload Me
    
'    Else
'        Me.lvPlatos.ListItems.Remove Me.lvPlatos.SelectedItem.Index
'        Me.lblTot.Caption = FormatCurrency(sumatoria, 2)
'        Me.lblItems.Caption = "Items: " & Me.lvPlatos.ListItems.count
       
    End If
    
     'ACTUALIZA LA COMANDA A FACTURADA Y LIBERA LA MESA CUANDO SE EXTORNA UN PLATO
    oCmdEjec.CommandText = "SP_PEDIDO_FACTURADO"
    LimpiaParametros oCmdEjec
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , Me.lblNumero.Caption)
        
    oCmdEjec.Execute

    Pub_ConnAdo.CommitTrans
'    If cCant = 1 Then
'    Unload Me
'    End If
    
    Exit Sub

eli:
    Pub_ConnAdo.RollbackTrans
    MsgBox Err.Description
End Sub

Private Sub extorna(cIDmotivo As Integer, cMOTIVO As String, cUSUARIO As String, cCodPlato As Double, cNumSec As Integer, cCantidad As Integer)
'Private Sub extorna()
Dim cCant As Integer

    cCant = Me.lvPlatos.ListItems.count
    
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
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numsec", adInteger, adParamInput, , Me.lvPlatos.SelectedItem.SubItems(6))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numsec", adInteger, adParamInput, , cNumSec)
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
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodArt", adBigInt, adParamInput, , Me.lvPlatos.SelectedItem.Tag)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodArt", adBigInt, adParamInput, , cCodPlato)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@cp", adBigInt, adParamInput, , Me.lvPlatos.SelectedItem.SubItems(3))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@cp", adBigInt, adParamInput, , cCantidad)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ser", adChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@nro", adInteger, adParamInput, , Me.lblNumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@mesa", adChar, adParamInput, 10, vMesa)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@tipo", adBoolean, adParamInput, , 0)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumSec", adInteger, adParamInput, , 0)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MaxNumFac", adInteger, adParamOutput, , 3)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MaxNumOper", adInteger, adParamOutput, , 3)
    
    oCmdEjec.Execute

    If Me.lvPlatos.ListItems.count = 1 Then
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SpLiberarMesa"
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodMesa", adVarChar, adParamInput, 10, vMesa)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodZon", adInteger, adParamInput, , vCodZona)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adVarChar, adParamInput, 3, Me.lblSerie.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , Me.lblNumero.Caption)
        oCmdEjec.Execute

'        Unload Me
'    Else
'        Me.lvPlatos.ListItems.Remove Me.lvPlatos.SelectedItem.Index
'        Me.lblTot.Caption = FormatCurrency(sumatoria, 2)
'        Me.lblItems.Caption = "Items: " & Me.lvPlatos.ListItems.count
    End If

    'ACTUALIZA LA COMANDA A FACTURADA Y LIBERA LA MESA CUANDO SE EXTORNA UN PLATO
    oCmdEjec.CommandText = "SP_PEDIDO_FACTURADO"
    LimpiaParametros oCmdEjec
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , Me.lblNumero.Caption)
    
    oCmdEjec.Execute
     
    Pub_ConnAdo.CommitTrans

    Exit Sub

ext:
    Pub_ConnAdo.RollbackTrans
    MsgBox Err.Description
End Sub

''''''Private Sub MoverPlato(Arriba As Boolean)
''''''If Me.lvPlatos.SelectedItem Is Nothing Then Exit Sub
''''''Dim itmX As ListItem
''''''Dim Plato As String
''''''Dim Cantidad As Double
''''''Dim PRECIO As Double
''''''Dim Importe As Double
''''''Dim CodPlato As Integer
''''''Dim DETALLE As String
''''''
''''''Dim lngIndex As Long
'''''''Obteniendo los valores seleccionados
''''''Plato = Me.lvPlatos.SelectedItem
''''''DETALLE = Me.lvPlatos.SelectedItem.SubItems(1)
''''''Cantidad = Me.lvPlatos.SelectedItem.SubItems(2)
''''''PRECIO = Me.lvPlatos.SelectedItem.SubItems(3)
''''''Importe = Me.lvPlatos.SelectedItem.SubItems(4)
''''''CodPlato = Me.lvPlatos.SelectedItem.Tag
''''''lngIndex = Me.lvPlatos.SelectedItem.Index
''''''
''''''If Arriba Then
''''''    If lngIndex = 1 Then Exit Sub
''''''    Me.lvPlatos.ListItems.Remove lngIndex
''''''    Set itmX = Me.lvPlatos.ListItems.Add(lngIndex - 1, , Plato, Me.ilComanda.ListImages(1).Key, Me.ilComanda.ListImages(1).Key)
''''''    itmX.Tag = CodPlato
''''''    itmX.SubItems(1) = DETALLE
''''''    itmX.SubItems(2) = FormatNumber(Cantidad, 2)
''''''    itmX.SubItems(3) = FormatNumber(PRECIO, 2)
''''''    itmX.SubItems(4) = FormatNumber(Importe, 2)
''''''    Me.lvPlatos.ListItems(lngIndex - 1).Selected = True
''''''    Me.lblTot.Caption = FormatCurrency(sumatoria, 2)
''''''Else
''''''    If lngIndex = Me.lvPlatos.ListItems.count Then Exit Sub
''''''    Me.lvPlatos.ListItems.Remove lngIndex
''''''    Set itmX = Me.lvPlatos.ListItems.Add(lngIndex + 1, , Plato, Me.ilComanda.ListImages(1).Key, Me.ilComanda.ListImages(1).Key)
''''''    itmX.Tag = CodPlato
''''''    itmX.SubItems(1) = DETALLE
''''''    itmX.SubItems(2) = FormatNumber(Cantidad, 2)
''''''    itmX.SubItems(3) = FormatNumber(PRECIO, 2)
''''''    itmX.SubItems(4) = FormatNumber(Importe, 2)
''''''    Me.lvPlatos.ListItems(lngIndex + 1).Selected = True
''''''    Me.lblTot.Caption = FormatCurrency(sumatoria, 2)
''''''End If
''''''
''''''
'''''''If Me.lvPlatos.ListItems.Count = lngIndex Then Exit Sub
'''''''Me.lvPlatos.ListItems.Remove (lngIndex)
'''''''If lngIndex = Me.lvPlatos.ListItems.Count + 1 Then
'''''''lngIndex = 0
'''''''End If
'''''''Set itmX = Me.lvPlatos.ListItems.Add(lngIndex + 1, , strCol1)
'''''''itmX.SubItems(1) = strCol2
'''''''
'''''''Me.lvPlatos.ListItems(lngIndex + 1).Selected = True
''''''End Sub

Private Function CrearEstructuraXML() As String
Dim vCadena As String
Dim itemP As ListItem
vCadena = "<r>"
For Each itemP In Me.lvPlatos.ListItems
    vCadena = vCadena & "<d "
    vCadena = vCadena & "cp=""" & itemP.Tag & """ "
    vCadena = vCadena & "d=""" & itemP.SubItems(1) & """ "
    vCadena = vCadena & "cant=""" & itemP.SubItems(2) & """ "
    vCadena = vCadena & "pre=""" & itemP.SubItems(3) & """ "
    vCadena = vCadena & "imp=""" & itemP.SubItems(4) & """ "
    vCadena = vCadena & "/>"
Next
vCadena = vCadena & "</r>"

CrearEstructuraXML = vCadena
End Function

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

Private Sub CargarPlatos()
oCmdEjec.CommandText = "SpListarPlatos"
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
valor = vPlato / 18
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
    If c <= 4 Then '1 fila
        If c = 1 Then
            vIniLeft = vIniLeft + Me.cmdPlatoAnt.Width
        Else
            vIniLeft = vIniLeft + Me.cmdPlato(i - i).Width
        End If
'        Else: viniLeft = viniLeft + 970
'        End If
    ElseIf c <= 9 Then '2º Fila
        'viniLeft = 30
        If c = 5 Then
            vIniLeft = 30
            vIniTop = vIniTop + Me.cmdPlatoAnt.Height
        Else: vIniLeft = vIniLeft + Me.cmdPlato(i - 1).Width
        End If
    ElseIf c <= 14 Then '3º Fila
        If c = 10 Then
            vIniTop = vIniTop + Me.cmdPlato(4).Height
            vIniLeft = 30
        Else: vIniLeft = vIniLeft + Me.cmdPlato(i - 1).Width
        End If
    Else '4º y ultima fila
        If c = 15 Then
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
    If c = 18 Then
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
 If Not IsNull(oRS!color) Then Me.cmdSubFam(i).BackColor = Trim(oRS!color)
    
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

Private Sub ConfiguraLV()
With Me.lvPlatos
'    .Gridlines = True
'    .LabelEdit = lvwManual
'    .FullRowSelect = True
'    .View = lvwReport
'
'    .HideSelection = False
'    .ColumnHeaders.Add , , "Descripción", 4000
'    .ColumnHeaders.Add , , "Cta.", 700
'    .ColumnHeaders.Add , , "Detalle", 1000
'    .ColumnHeaders.Add , , "Cant.", 800, 1
'    .ColumnHeaders.Add , , "Precio", 800, 1
'    .ColumnHeaders.Add , , "Importe", 1000, 1
'    .ColumnHeaders.Add , , "numsec", 0
'    .ColumnHeaders.Add , , "Atend", 0
'    .ColumnHeaders.Add , , "allnumfac", 0
'    .ColumnHeaders.Add , , "apro", 0
'    .ColumnHeaders.Add , , "numfac", 0
'    .MultiSelect = True

.Gridlines = True
    .LabelEdit = lvwManual
    .FullRowSelect = True
    .View = lvwReport
    
    .HideSelection = False
    .ColumnHeaders.Add , , "Descripción", 4000
    .ColumnHeaders.Add , , "Cta.", 700
    .ColumnHeaders.Add , , "Detalle", 1000
    .ColumnHeaders.Add , , "Cant.", 800, 1
    .ColumnHeaders.Add , , "Precio", 800, 1
    .ColumnHeaders.Add , , "Importe", 1000, 1
    .ColumnHeaders.Add , , "numsec", 0
    .ColumnHeaders.Add , , "Atend", 0
    .ColumnHeaders.Add , , "allnumfac", 0
    .ColumnHeaders.Add , , "apro", 0
    .ColumnHeaders.Add , , "numfac", 0
    .ColumnHeaders.Add , , "cantado", 0
    .ColumnHeaders.Add , , "Enviar", 1200
    .ColumnHeaders.Add , , "FAM", 200
    .MultiSelect = True
End With
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
        If Me.lvPlatos.SelectedItem.SubItems(9) = 1 Then 'Or val(Me.lvPlatos.SelectedItem.SubItems(7)) <> 0 Then
            'If MsgBox("no se puede MODIFICAR el plato, ya fue despachado" & vbCrLf & "¿desea ingresar la clave?", vbQuestion + vbYesNo, NombreProyecto) = vbYes Then
            MsgBox ("no se puede MODIFICAR el plato, ya fue despachado")
           
        Else
    
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

    Exit Sub

elimina:
    MsgBox Err.Description

End Sub

Private Sub cmdCarac_Click()

    If Me.lvPlatos.ListItems.count = 0 Then Exit Sub
    frmComandaProdCaracteristicas.gIDproducto = Me.lvPlatos.SelectedItem.Tag
    frmComandaProdCaracteristicas.gNUMFAC = Me.lblNumero.Caption
    frmComandaProdCaracteristicas.gNUMSER = Me.lblSerie.Caption
    frmComandaProdCaracteristicas.gNUMSEC = Me.lvPlatos.SelectedItem.SubItems(6)
    frmComandaProdCaracteristicas.Show vbModal
End Sub

Private Sub cmdCerrar_Click()

Unload Me
End Sub

Private Sub cmdCta_Click()
If Not Me.lvPlatos.SelectedItem Is Nothing Then
    Dim Dato As String
    frmDetalle.EsDetalle = True
    frmDetalle.Show vbModal
    If frmDetalle.vSelec Then


        Dim i As Integer
        For i = 1 To Me.lvPlatos.ListItems.count



        If Me.lvPlatos.ListItems(i).Selected Then
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SpActualizarDetallePlato"
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adChar, adParamInput, 3, Me.lblSerie.Caption)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adDouble, adParamInput, , CDbl(Me.lblNumero.Caption))
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numsec", adInteger, adParamInput, , CInt(Me.lvPlatos.ListItems(i).SubItems(6)))
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@det", adVarChar, adParamInput, 50, frmDetalle.vDetalle)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ESDETALLE", adBoolean, adParamInput, , 1)
            oCmdEjec.Execute

            Me.lvPlatos.ListItems(i).SubItems(1) = frmDetalle.vDetalle
        End If
        Next



    End If
End If
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

                CargarComanda LK_CODCIA, vMesa
            End If

        Else
            MsgBox "Clave incorrecta", vbCritical, NombreProyecto
        End If
    End If

End Sub

Private Sub cmdDetalle_Click()
If Not Me.lvPlatos.SelectedItem Is Nothing Then
    Dim Dato As String
    frmDetalle.Show vbModal
    frmDetalle.EsDetalle = False
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

Private Sub cmdEliminar_Click()

'    On Error GoTo elimina
'
'    If Not Me.lvPlatos.SelectedItem Is Nothing Then
'        If Me.lvPlatos.SelectedItem.SubItems(9) = 1 Or vEstado = "E" Then  'Or val(Me.lvPlatos.SelectedItem.SubItems(6)) <> 0 Then
'            If MsgBox("no se puede eliminar el plato, ya fue despachado o la Mesa está EN CUENTA" & vbCrLf & "¿desea ingresar la clave?", vbQuestion + vbYesNo, NombreProyecto) = vbYes Then
'                frmClaveCaja.Show vbModal
'
'                If frmClaveCaja.vAceptar Then
'
'                    Dim VS As String
'
'                    If VerificaPass(frmClaveCaja.vUSUARIO, frmClaveCaja.vClave, VS) Then
'                        extorna
'                    Else
'                        MsgBox "Clave incorrecta", vbCritical, NombreProyecto
'                    End If
'                End If
'            End If
'
'        Else
'            elimina
'        End If
'    End If
'
'    Exit Sub
'
'elimina:
'    MsgBox Err.Description

    On Error GoTo elimina

    'Dim VS As String
    Dim cElimina As Boolean

    cElimina = False

    Dim xRP As Boolean

    xRP = False

    If Not Me.lvPlatos.SelectedItem Is Nothing Then
    
        For Each xITEM In Me.lvPlatos.ListItems

            If xITEM.Selected Then
                'VERIFICA TIEMPO
                LimpiaParametros oCmdEjec
                oCmdEjec.CommandText = "SP_VERIFICA_TIEMPO_REGISTRO"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adChar, adParamInput, 3, Me.lblSerie.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adDouble, adParamInput, , CDbl(Me.lblNumero.Caption))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numsec", adInteger, adParamInput, , CInt(Me.lvPlatos.SelectedItem.SubItems(6)))

                Dim ORSd As ADODB.Recordset

                Set ORSd = oCmdEjec.Execute

                Dim xdato As Boolean
        
                If Not ORSd.EOF Then
                    xdato = ORSd!Dato
                End If
            
                If xITEM.SubItems(9) = 1 Or xITEM.SubItems(7) <> 0 Or vEstado = "E" Or val(xITEM.SubItems(11)) = 1 Or xdato = False Then
                    xRP = True

                    Exit For

                End If
            End If

        Next

    End If
    
    If xRP Then
        If MsgBox("no se puede eliminar el plato, ya fue despachado o la Mesa está EN CUENTA, o el Item está en Preparación." & vbCrLf & "¿desea ingresar la clave?", vbQuestion + vbYesNo, NombreProyecto) = vbYes Then
            frmClaveCaja.Show vbModal
    
            If frmClaveCaja.vAceptar Then
    
                Dim vS As String
    
                If VerificaPass(frmClaveCaja.vUSUARIO, frmClaveCaja.vClave, vS) Then
                    frmComandaMotivoElimina.Show vbModal
    
                    If frmComandaMotivoElimina.gAcepta Then

                        For Each xITEM In Me.lvPlatos.ListItems

                            If xITEM.Selected Then
                                extorna frmComandaMotivoElimina.gIDmotivo, frmComandaMotivoElimina.gMOTIVO, frmClaveCaja.vUSUARIO, xITEM.Tag, CDbl(xITEM.SubItems(6)), CDbl(xITEM.SubItems(3))
                                cElimina = True
                            End If

                        Next
                      
                    End If

                Else
                    MsgBox "Clave incorrecta", vbCritical, NombreProyecto
                End If
            End If
        End If

    Else

        For Each xITEM In Me.lvPlatos.ListItems

            If xITEM.Selected Then
                elimina xITEM.SubItems(6), xITEM.Tag, xITEM.SubItems(3)
                cElimina = True
            End If

        Next
       
    End If
            
    If cElimina Then

        Dim i As Integer

        For i = Me.lvPlatos.ListItems.count To 1 Step -1

            If Me.lvPlatos.ListItems(i).Selected Then
                Me.lvPlatos.ListItems.Remove Me.lvPlatos.ListItems(i).index
            End If
 
        Next

        Me.lblTot.Caption = FormatCurrency(sumatoria, 2)
        Me.lblItems.Caption = "Items: " & Me.lvPlatos.ListItems.count
    End If

    If Me.lvPlatos.ListItems.count = 0 Then
        Unload Me
    Else
        Me.lvPlatos.ListItems(Me.lvPlatos.ListItems.count).Selected = True
    End If

    Exit Sub

elimina:
    MsgBox Err.Description

End Sub

Private Sub cmdFactura_Click()

    If Me.lvPlatos.ListItems.count > 0 Then
        frmFacComanda.vTOTAL = Me.lblTot.Caption
        frmFacComanda.vMesa = "0"
        frmFacComanda.vNroCom = Me.lblNumero.Caption
        frmFacComanda.vSerCom = Me.lblSerie.Caption
        frmFacComanda.vCodMoz = Me.lblMozo.Tag
        frmFacComanda.xMostrador = True
        frmFacComanda.Show vbModal

        If frmFacComanda.vAcepta Then
            'Unload Me
            Me.lvPlatos.ListItems.Clear
            Me.lblSerie.Caption = ""
            Me.lblNumero.Caption = ""
            Me.lblTot.Caption = "S/. 0.00"
            Me.lblItems.Caption = "Items: 0"
            Me.lblCliente.Caption = ""
            VNuevo = True
        End If
    End If

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

Private Sub cmdGuardar_Click()
'If Me.lvPlatos.ListItems.Count = 0 Then
'    MsgBox "Debe Ingresar Platos en la Comanda", vbCritical, NombreProyecto
'    Exit Sub
'End If
'If Len(Trim(Me.lblMozo.Caption)) = 0 Then
'    MsgBox "Debe elegir el Mozo", vbCritical, NombreProyecto
'    Exit Sub
'End If
'On Error GoTo Graba
'LimpiaParametros oCmdEjec
'Dim xPedido As String
'Dim NumSer As String
'Dim NumFac As Integer
'xPedido = CrearEstructuraXML
'
'oCmdEjec.CommandType = adCmdStoredProc
'If VNuevo Then 'nueva comanda
'    With oCmdEjec
'        .CommandText = "SpRegistrarComanda"
'        .Parameters.Append .CreateParameter("@CodCia", adChar, adParamInput, 2, lk_codcia)
'        .Parameters.Append .CreateParameter("@Usuario", adVarChar, adParamInput, 10, lk_codusu)
'        .Parameters.Append .CreateParameter("@CodMesa", adVarChar, adParamInput, 10, vMesa)
' '       .Parameters.Append .CreateParameter("@Total", adCurrency, adParamInput, , Val(Me.lblTot.Tag))
'        .Parameters.Append .CreateParameter("@Mozo", adInteger, adParamInput, , CInt(Me.lblMozo.Tag))
'        .Parameters.Append .CreateParameter("@xPedido", adVarChar, adParamInput, 4000, xPedido)
'        .Parameters.Append .CreateParameter("@NumSer", adChar, adParamOutput, 3, NumSer)
'        .Parameters.Append .CreateParameter("@NumFac", adInteger, adParamOutput, , NumFac)
'    End With
'    oCmdEjec.Execute
'    Me.lblSerie.Caption = oCmdEjec.Parameters("@NumSer").Value
'    Me.lblNumero.Caption = oCmdEjec.Parameters("@NumFac").Value
'    VNuevo = False
'    'Cambiando imagen de Mesa
'
'    For I = 1 To frmDisMesas.imgMesa.Count - 1
'        If frmDisMesas.lblNomMesa(I).Tag = vMesa Then
'            frmDisMesas.imgMesa(I).Picture = frmDisMesas.ilMesas.ListImages(4).Picture
'            frmDisMesas.imgMesa(I).ToolTipText = "Mesa Ocupada"
'            frmDisMesas.imgMesa(I).Tag = "O"
'        End If
'    Next
'
'    'oCmdEjec
'Else 'Modifica comanda
'     With oCmdEjec
'        .CommandText = "SpModificarComanda"
'        .Parameters.Append .CreateParameter("@CodCia", adChar, adParamInput, 2, lk_codcia)
'        .Parameters.Append .CreateParameter("@Usuario", adVarChar, adParamInput, 10, lk_codusu)
'        .Parameters.Append .CreateParameter("@CodMesa", adVarChar, adParamInput, 10, vMesa)
' '       .Parameters.Append .CreateParameter("@Total", adCurrency, adParamInput, , Val(Me.lblTot.Tag))
'        .Parameters.Append .CreateParameter("@Mozo", adInteger, adParamInput, , CInt(Me.lblMozo.Tag))
'        .Parameters.Append .CreateParameter("@xPedido", adVarChar, adParamInput, 4000, xPedido)
'        .Parameters.Append .CreateParameter("@NumSer", adChar, adParamInput, 3, Me.lblSerie.Caption)
'        .Parameters.Append .CreateParameter("@NumFac", adInteger, adParamInput, , Me.lblNumero.Caption)
'        If Len(Trim(vMesaAnt)) <> 0 Then
'            .Parameters.Append .CreateParameter("@CodMesaAnt", adVarChar, adParamInput, 10, vMesaAnt)
'            For I = 1 To frmDisMesas.imgMesa.Count - 1
'                If frmDisMesas.lblNomMesa(I).Tag = vMesa Then
'                    frmDisMesas.imgMesa(I).Picture = frmDisMesas.ilMesas.ListImages(4).Picture
'                    frmDisMesas.imgMesa(I).ToolTipText = "Mesa Ocupada"
'                    frmDisMesas.imgMesa(I).Tag = "O"
'                ElseIf frmDisMesas.lblNomMesa(I).Tag = vMesaAnt Then
'                    frmDisMesas.imgMesa(I).Picture = frmDisMesas.ilMesas.ListImages(1).Picture
'                    frmDisMesas.imgMesa(I).ToolTipText = "Mesa Libre"
'                    frmDisMesas.imgMesa(I).Tag = "L"
'                End If
'            Next
'            vMesaAnt = ""
'        End If
'    End With
'    oCmdEjec.Execute
'End If
'If MsgBox("¿Desea imprimir la Comanda?", vbQuestion + vbYesNo, NombreProyecto) = vbYes Then
'    cmdPrint_Click
'End If
'Exit Sub
'Graba:
' MsgBox Err.Description, vbCritical, NombreProyecto
End Sub


Private Sub cmdLimpiar_Click()
Me.lblTexto.Caption = ""
End Sub






Private Sub cmdNum_Click(index As Integer)
Me.lblTexto.Caption = Me.lblTexto.Caption & Me.cmdnum(index).Caption
End Sub

Private Sub cmdPlato_Click(index As Integer)
            
    Dim c As Integer

    For c = 1 To Me.lvPlatos.ListItems.count
        Me.lvPlatos.ListItems(c).Selected = False
    Next

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

    'If vcero Then Exit Sub

    If VNuevo Then
        If Me.lvPlatos.ListItems.count = 0 Then
            'obteniendo precio
            oRsPlatos.Filter = "Codigo = '" & Me.cmdPlato(index).Tag & "'"
       
            If AgregaPlato(Me.cmdPlato(index).Tag, 1, FormatNumber(oRsPlatos!PRECIO, 2), oRsPlatos!PRECIO, "", "", 0, Me.lblCliente.Caption, IIf(Len(Trim(Me.lblComensales.Caption)) = 0, 0, Me.lblComensales.Caption)) Then
        
                With Me.lvPlatos.ListItems.Add(, , Me.cmdPlato(index).Caption, Me.ilComanda.ListImages.Item(1).key, Me.ilComanda.ListImages.Item(1).key)
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
                    'ItemP.Checked = True
                    oRsPlatos.Filter = ""
                    oRsPlatos.MoveFirst

                End With
                
                vEstado = "O"
                adicionarBolsa Me.cmdPlato(index).Tag, Me.lblSerie.Caption, Me.lblNumero.Caption, Me.lblCliente.Caption
            End If
        End If

    Else

        Dim DD As Integer

        oRsPlatos.Filter = "Codigo = '" & Me.cmdPlato(index).Tag & "'"

        If AgregaPlato(Me.cmdPlato(index).Tag, 1, FormatNumber(oRsPlatos!PRECIO, 2), oRsPlatos!PRECIO, "", Me.lblSerie.Caption, Me.lblNumero.Caption, Me.lblCliente.Caption, IIf(Len(Trim(Me.lblComensales.Caption)) = 0, 0, frmComanda2.lblComensales.Caption), DD) Then
    
            With Me.lvPlatos.ListItems.Add(, , Me.cmdPlato(index).Caption, Me.ilComanda.ListImages.Item(1).key, Me.ilComanda.ListImages.Item(1).key)
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
        End If
    End If

    If Me.lvPlatos.ListItems.count <> 0 Then
        'Me.lvPlatos.SelectedItem = Nothing
        Me.lblTot.Caption = FormatCurrency(sumatoria, 2)
        Me.lblItems.Caption = "Items: " & Me.lvPlatos.ListItems.count
        Me.lvPlatos.ListItems(Me.lvPlatos.ListItems.count).Selected = True
    End If

    'Dim I As Integer
    'prueba

    CargarComanda LK_CODCIA, vMesa
    'aqui

    'Dim C As Integer
    For c = 1 To Me.lvPlatos.ListItems.count
        'If Me.lvPlatos.ListItems(c).Checked Then
        Me.lvPlatos.ListItems(c).Selected = False
        'End If
    Next

    Me.lvPlatos.ListItems(Me.lvPlatos.ListItems.count).Selected = True
End Sub

Private Sub cmdPlatoAnt_Click()
Dim ini, fin, f As Integer
If vPagActFam = 2 Then
    ini = 1
    fin = 18
ElseIf vPagActPla = 1 Then
    Exit Sub
Else
    FF = vPagActPla - 1
    ini = (18 * FF) - 17
    fin = 18 * FF
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
    fin = 18
ElseIf vPagActPla = vPagTotPla Then
    Exit Sub
Else
    ini = (18 * vPagActPla) - 17
    fin = 18 * vPagActPla
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

Private Sub CambiaPrecio()
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
 
End Sub

Private Sub cmdPorcion_Click()

    If Me.lvPlatos.ListItems.count = 0 Then Exit Sub

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_PEDIDO_VALIDAPORCION"

    Dim orsP As ADODB.Recordset

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , Me.lvPlatos.SelectedItem.Tag)

    Set orsP = oCmdEjec.Execute

    If CBool(orsP!porc) Then

        On Error GoTo Porcion

        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SP_PEDIDO_CONVERTIR_PORCION"

        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, 2, Me.lblNumero.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adVarChar, adParamInput, 3, Me.lblSerie.Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSEC", adInteger, adParamInput, , Me.lvPlatos.SelectedItem.SubItems(6))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , Me.lvPlatos.SelectedItem.Tag)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PRECIO", adDouble, adParamInput, , orsP!PPORC)
        oCmdEjec.Execute

        Me.lvPlatos.SelectedItem.SubItems(4) = Format(orsP!PPORC, "#####.#0")
        Me.lvPlatos.SelectedItem.SubItems(3) = 0.5
        Me.lvPlatos.SelectedItem.SubItems(5) = Format(orsP!PPORC * 0.5, "#####.#0")

        Me.lblTot.Caption = FormatCurrency(sumatoria, 2)

        Exit Sub

Porcion:
        MsgBox Err.Description, vbCritical, Pub_Titulo

    Else
        MsgBox "El Producto no permite porcion.", vbCritical, Pub_Titulo
    End If

End Sub

Private Sub cmdPrecio_Click()
' If Me.lvPlatos.SelectedItem.SubItems(6) <> 1 Or Me.lvPlatos.SelectedItem.SubItems(5) <> 0 Then
' If Me.lvPlatos.SelectedItem.SubItems(8) = 1 Then  'Or val(Me.lvPlatos.SelectedItem.SubItems(5)) <> 0 Then
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
    'Else
    '    CambiaPrecio
    'End If

    

End Sub

Private Sub cmdPreCuenta_Click()
On Error GoTo printe
If Len(Trim(Me.lblNumero.Caption)) = 0 And Len(Trim(Me.lblNumero.Caption)) = 0 Then
MsgBox "No hay nada que imprimir", vbCritical, NombreProyecto
Else

Dim ORSSepara As ADODB.Recordset

LimpiaParametros oCmdEjec
oCmdEjec.CommandType = adCmdStoredProc
oCmdEjec.CommandText = "SpSeparaCuentas"

oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCLIE", adVarChar, adParamInput, 15, "0")
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumSer", adChar, adParamInput, 3, Me.lblSerie.Caption)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumFac", adBigInt, adParamInput, , Me.lblNumero.Caption)


Set ORSSepara = oCmdEjec.Execute

Do While Not ORSSepara.EOF


Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
Dim crParamDef As CRAXDRT.ParameterFieldDefinition
Dim objCrystal As New CRAXDRT.APPLICATION

Dim RutaReporte As String
RutaReporte = "C:\Admin\Nordi\Comanda2.rpt"

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
oCmdEjec.CommandText = "SpPrintComanda2"
'oCmdEjec.CommandText = "SpPrintComanda"

Dim rsd As ADODB.Recordset

Dim vdata As String
Dim vnumsec As String
vdata = ""
Dim c As Integer
For c = 1 To Me.lvPlatos.ListItems.count
    'If Me.lvPlatos.ListItems(c).Checked Then
        vdata = vdata & Me.lvPlatos.ListItems(c).Tag & ","
        vnumsec = vnumsec & Me.lvPlatos.ListItems(c).SubItems(6) & ","
    'End If
Next

oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumSer", adChar, adParamInput, 3, Me.lblSerie.Caption)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NumFac", adDouble, adParamInput, , Me.lblNumero.Caption)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@xdet", adVarChar, adParamInput, 4000, vdata)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@xnumsec", adVarChar, adParamInput, 4000, vnumsec)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@precuenta", adBoolean, adParamInput, , 1) 'JULIO 11-01-2011
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CTA", adChar, adParamInput, 1, ORSSepara!cuenta)

Set rsd = oCmdEjec.Execute

'LimpiaParametros oCmdEjec
'oCmdEjec.CommandText = "SpMesaEnCuenta"
'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("CodMesa", adVarChar, adParamInput, 10, vMesa)
'oCmdEjec.Execute
'
'  For i = 1 To frmDisMesas.imgMesa.count - 1
'        If frmDisMesas.lblNomMesa(i).Tag = vMesa Then
'            frmDisMesas.imgMesa(i).Picture = frmDisMesas.ilMesas.ListImages(2).Picture
'            frmDisMesas.imgMesa(i).ToolTipText = "Mesa En Cuenta"
'            frmDisMesas.imgMesa(i).Tag = "E"
'        End If
'    Next
'    vEstado = "E"



VReporte.DataBase.SetDataSource rsd, 3, 1
'frmprint.CRViewer1.ReportSource = VReporte
'frmprint.CRViewer1.ViewReport
VReporte.PrintOut , 1, , 1, 1
Set objCrystal = Nothing
Set VReporte = Nothing






    ORSSepara.MoveNext
Loop


Exit Sub
printe:
MostrarErrores Err
End If

End Sub

Private Sub cmdPrint_Click()
If Len(Trim(Me.lblNumero.Caption)) = 0 And Len(Trim(Me.lblNumero.Caption)) = 0 Then
MsgBox "No hay nada que imprimir", vbCritical, NombreProyecto
Else

Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
Dim crParamDef As CRAXDRT.ParameterFieldDefinition
Dim objCrystal As New CRAXDRT.APPLICATION

Dim RutaReporte As String
RutaReporte = "C:\Admin\Nordi\Comanda1.rpt"

'Verificar platos enviados para mensaje
Dim cat As Integer
Dim Mensaje As String
Dim mATRIZ() As Integer
Dim ss As Integer

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
            crParamDef.AddCurrentValue str(vPlato)
            Case "Mensaje"
            crParamDef.AddCurrentValue Mensaje
    End Select
Next

On Error GoTo printe
LimpiaParametros oCmdEjec
oCmdEjec.CommandType = adCmdStoredProc
oCmdEjec.CommandText = "SpPrintComanda2"
'oCmdEjec.CommandText = "SpPrintComanda"

Dim rsd As ADODB.Recordset

Dim vdata As String
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

        'Exit Sub

        '        'COCINA
        '        rsd.Filter = "PED_FAMILIA = 1 OR PED_FAMILIA = 2"
        '
        '        Dim dd As ADODB.Recordset
        '
        '        If Not rsd.EOF Then
        '
        '            VReporte.Database.SetDataSource rsd, 3, 1 'lleno el objeto reporte
        '            'VReporte.SelectPrinter Printer.DriverName, "\\CAJA01\Cocina", Printer.Port
        '            ' VReporte.SelectPrinter Printer.DriverName, "Cocina", Printer.Port
        '            ' VReporte.SelectPrinter Printer.DriverName, "\\Mozos\Cocina", Printer.Port
        '            ' VReporte.SelectPrinter Printer.DriverName, "\\Caja2\Cocina1", Printer.Port
        '            VReporte.SelectPrinter Printer.DriverName, "\\Cocina\Cocina", Printer.Port
        '            'VReporte.SelectPrinter Printer.DriverName, "Cocina", Printer.Port
        '            VReporte.PrintOut False, 1, , 1, 1
        '
        '            Set VReporte = Nothing
        '            Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
        '
        '            Set crParamDefs = VReporte.ParameterFields
        '
        '            For Each crParamDef In crParamDefs
        '
        '                Select Case crParamDef.ParameterFieldName
        '
        '                    Case "mesa"
        '                        crParamDef.AddCurrentValue str(vPlato)
        '
        '                    Case "Mensaje"
        '                        crParamDef.AddCurrentValue Mensaje
        '                End Select
        '
        '            Next
        '
        '        End If
        '
        '        rsd.Filter = "PED_FAMILIA=3"
        '
        '        If Not rsd.EOF Then
        '            VReporte.Database.SetDataSource rsd, 3, 1 'lleno el objeto reporte
        '            'VReporte.SelectPrinter Printer.DriverName, "\\laptop\doPDF v6", Printer.Port
        '            ' VReporte.SelectPrinter Printer.DriverName, "\\CAJA01\jugos", Printer.Port 'doPDF v6
        '            'VReporte.SelectPrinter Printer.DriverName, "jugos", Printer.Port
        '            VReporte.SelectPrinter Printer.DriverName, "\\Punto2\Bar", Printer.Port
        '            ' VReporte.SelectPrinter Printer.DriverName, "\\VENTAS1\Bar", Printer.Port
        '            '   VReporte.SelectPrinter Printer.DriverName, "Bar", Printer.Port
        '            VReporte.PrintOut False, 1, , 1, 1
        '
        '            Set VReporte = Nothing
        '            Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
        '
        '            Set crParamDefs = VReporte.ParameterFields
        '
        '            For Each crParamDef In crParamDefs
        '
        '                Select Case crParamDef.ParameterFieldName
        '
        '                    Case "mesa"
        '                        crParamDef.AddCurrentValue str(vPlato)
        '
        '                    Case "Mensaje"
        '                        crParamDef.AddCurrentValue Mensaje
        '                End Select
        '
        '            Next
        '
        '        End If
        '
        '        'rsd.Filter = "PED_FAMILIA=3"
        '        'If Not rsd.EOF Then
        '        '    VReporte.Database.SetDataSource rsd, 3, 1 'lleno el objeto reporte
        '        '    'VReporte.SelectPrinter Printer.DriverName, "\\laptop\doPDF v6", Printer.Port
        '        '   ' VReporte.SelectPrinter Printer.DriverName, "\\Digitacion\Bar", Printer.Port 'doPDF v6
        '        '    VReporte.SelectPrinter Printer.DriverName, "Bar01", Printer.Port
        '        '   ' VReporte.SelectPrinter Printer.DriverName, "\\Caja\Bar01", Printer.Port   'SR BEFE
        '        '    VReporte.PrintOut ' , 1, , 1, 1
        '        'End If
        '
        '        'Set VReporte = Nothing
        '        'Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
        '
        '        'Set crParamDefs = VReporte.ParameterFields
        '        'For Each crParamDef In crParamDefs
        '        '    Select Case crParamDef.ParameterFieldName
        '        '        Case "mesa"
        '        '            crParamDef.AddCurrentValue str(vPlato)
        '        '            Case "Mensaje"
        '        '            crParamDef.AddCurrentValue Mensaje
        '        '    End Select
        '        'Next
        '
        '        'rsd.Filter = "PED_FAMILIA=4"
        '        'If Not rsd.EOF Then
        '        '    VReporte.Database.SetDataSource rsd, 3, 1 'lleno el objeto reporte
        '        '   ' VReporte.SelectPrinter Printer.DriverName, "\\laptop\doPDF v6", Printer.Port
        '        '    'VReporte.SelectPrinter Printer.DriverName, "\\CAJA\Star SP542 Line Mode Printer with Status Monitor", Printer.Port 'doPDF v6
        '        '    VReporte.SelectPrinter Printer.DriverName, "\\MOZOS2\Bar02", Printer.Port
        '        '   ' VReporte.SelectPrinter Printer.DriverName, "Bar02", Printer.Port     'SR BEFE
        '        '    VReporte.PrintOut 'false, 1, , 1, 1
        '        'End If
        '        'Set VReporte = Nothing
        '        'Set VReporte = objCrystal.OpenReport(RutaReporte, 1)'
        '
        '        'Set crParamDefs = VReporte.ParameterFields
        '        'For Each crParamDef In crParamDefs
        '        '    Select Case crParamDef.ParameterFieldName
        '        '        Case "mesa"
        '        '            crParamDef.AddCurrentValue str(vPlato)
        '        '            Case "Mensaje"
        '        '            crParamDef.AddCurrentValue Mensaje
        '        '    End Select
        '        'Next
        '
        '        rsd.Filter = "PED_FAMILIA=6"
        '
        '        If Not rsd.EOF Then
        '            VReporte.Database.SetDataSource rsd, 3, 1 'lleno el objeto reporte
        '            'VReporte.SelectPrinter Printer.DriverName, "\\laptop\doPDF v6", Printer.Port
        '            ' VReporte.SelectPrinter Printer.DriverName, "\\CAJA01\jugos", Printer.Port 'doPDF v6
        '            'VReporte.SelectPrinter Printer.DriverName, "jugos", Printer.Port
        '            'VReporte.SelectPrinter Printer.DriverName, "\\Cafetin-pc\cafetin", Printer.Port   'SR BEFE
        '            VReporte.SelectPrinter Printer.DriverName, "\\digitacion\Cocina", Printer.Port   'mochica
        '            ' VReporte.SelectPrinter Printer.DriverName, "Bar", Printer.Port
        '            VReporte.PrintOut False, 1, , 1, 1
        '
        '            Set VReporte = Nothing
        '            Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
        '
        '            Set crParamDefs = VReporte.ParameterFields
        '
        '            For Each crParamDef In crParamDefs
        '
        '                Select Case crParamDef.ParameterFieldName
        '
        '                    Case "mesa"
        '                        crParamDef.AddCurrentValue str(vPlato)
        '
        '                    Case "Mensaje"
        '                        crParamDef.AddCurrentValue Mensaje
        '                End Select
        '
        '            Next
        '
        '        End If
        '
        '        rsd.Filter = "PED_FAMILIA=7"
        '
        '        If Not rsd.EOF Then
        '            VReporte.Database.SetDataSource rsd, 3, 1 'lleno el objeto reporte
        '            'VReporte.SelectPrinter Printer.DriverName, "\\laptop\doPDF v6", Printer.Port
        '            ' VReporte.SelectPrinter Printer.DriverName, "\\CAJA01\jugos", Printer.Port 'doPDF v6
        '            'VReporte.SelectPrinter Printer.DriverName, "jugos", Printer.Port
        '            'VReporte.SelectPrinter Printer.DriverName, "\\Cafetin-pc\cafetin", Printer.Port   'SR BEFE
        '            VReporte.SelectPrinter Printer.DriverName, "\\caja\Bar", Printer.Port   'mochica
        '            ' VReporte.SelectPrinter Printer.DriverName, "Bar", Printer.Port
        '            VReporte.PrintOut False, 1, , 1, 1
        '
        '            Set VReporte = Nothing
        '            Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
        '
        '            Set crParamDefs = VReporte.ParameterFields
        '
        '            For Each crParamDef In crParamDefs
        '
        '                Select Case crParamDef.ParameterFieldName
        '
        '                    Case "mesa"
        '                        crParamDef.AddCurrentValue str(vPlato)
        '
        '                    Case "Mensaje"
        '                        crParamDef.AddCurrentValue Mensaje
        '                End Select
        '
        '            Next
        '
        '        End If

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
       ' Unload Me    'gts cierra comanda despues de imprimir
         If gDefecto Then
        '    Unload frmMainMesas
         End If
        Exit Sub

printe:
        MostrarErrores Err
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

Private Sub cmdSubir_Click()
'MoverPlato True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then cmdCerrar_Click
    If KeyCode = vbKeyF3 Then
        frmAsigna.Mostrador = True
        frmAsigna.Show vbModal
    End If

    If KeyCode = vbKeyF2 Then
        If Me.lvPlatos.ListItems.count <> 0 Then
            frmResumen.Caption = "Resumen de Productos: Comanda :" & Me.lblSerie.Caption & "-" & Me.lblNumero.Caption
            frmResumen.Show vbModal
        End If
    End If

    If KeyCode = vbKeyF5 Then
        frmComandaProductoSearch.gMostrador = True
        frmComandaProductoSearch.gDELIVERY = False
        frmComandaProductoSearch.Show vbModal
    End If
    
    If KeyCode = vbKeyF6 Then
    frmDetalle.Comensales = False
    frmDetalle.vDetalle = ""
    frmDetalle.Caption = "Ingrese el Nombre del Cliente"

    frmDetalle.txtMesa.Text = Me.lblCliente.Caption
    frmDetalle.ParaCliente = True
    frmDetalle.Show vbModal

    If frmDetalle.vSelec Then
        Me.lblCliente.Caption = frmDetalle.vDetalle

        If Not VNuevo Then
            ActualizaDatosComanda
        End If
    End If
    End If

End Sub

Private Sub Form_Load()
        'Me.WindowState = 0
CentrarFormulario MDIForm1, Me
    ' AGREGADO GTS PARA C¿VERIFICAR FECHA DEL DIA=========================================================
    SQ_OPER = 1
    PUB_CODCIA = LK_CODCIA
    LEER_PAR_LLAVE

    If par_llave!par_flag_cierre = 9 Then
        MsgBox "!!! Compañia ... Cerró Operaciones ... Llamar al Administrador ", 48, Pub_Titulo
        Unload Me

        'GoTo salirf
        Exit Sub

    Else
    End If

    If LK_FLAG_GRIFO <> "A" Then
        If par_llave!PAR_FECHA_DIA <> LK_FECHA_DIA Then
            MsgBox "!!!FECHA YA NO COINCIDE CON LA ACTUAL , OTRO USUARIO HA CERRADO EL DIA!!! SALGA Y REINICIE SU SISTEMA...", 48, Pub_Titulo

            End

            Unload Me
            'GoTo salirf
        End If
    End If

    ' AGREGADO GTS PARA C¿VERIFICAR FECHA DEL DIA=======================================================

    InhabilitarCerrar Me
    Me.lblTot.Caption = FormatCurrency("0.00", 2)
    vIniLeft = 30
    vIniTop = 120
    Me.Top = 120
    Me.Left = 30

    If LK_USU_STOCK = "A" Then
        cmdFactura.Enabled = True
    Else
        cmdFactura.Enabled = False
    End If
    
      If LK_USU_CUENTA = "A" Then
        cmdPreCuenta.Enabled = True
    Else
        cmdPreCuenta.Enabled = False
    End If
    
     If LK_USU_IMPRIME = "A" Then
        cmdPrint.Enabled = True
    Else
        cmdPrint.Enabled = False
    End If
 
    ConfiguraLV
    CargarFamilias
    CargarSubFamilias
    CargarPlatos

    CargarComanda LK_CODCIA, "0"
    
    If Me.lvPlatos.ListItems.count = 0 Then
        VNuevo = True
    Else
        VNuevo = False
    End If
    
    'Me.cmdMesa.Enabled = True

    Me.lblItems.Caption = "Items :" & Me.lvPlatos.ListItems.count

    If vPrimero Then
        Me.lvPlatos.CheckBoxes = False
    Else
        Me.lvPlatos.CheckBoxes = True
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
'frmDisMesas.CargarMesas frmDisMesas.VZONA
'frmDisMesas.WindowState = 2
End Sub

Private Sub lblCliente_Click()
    frmDetalle.Comensales = False
    frmDetalle.vDetalle = ""
    frmDetalle.Caption = "Ingrese el Nombre del Cliente"

    frmDetalle.txtMesa.Text = Me.lblCliente.Caption
    frmDetalle.ParaCliente = True
    frmDetalle.Show vbModal

    If frmDetalle.vSelec Then
        Me.lblCliente.Caption = frmDetalle.vDetalle

        If Not VNuevo Then
            ActualizaDatosComanda
        End If
    End If

End Sub

Private Sub ActualizaDatosComanda()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPACTUALIZARDATOSCOMANDA"
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adDouble, adParamInput, , Me.lblNumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CLIENTE", adVarChar, adParamInput, 120, Me.lblCliente.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@COMENSALES", adDouble, adParamInput, , IIf(Len(Trim(Me.lblComensales.Caption)) = 0, 0, Me.lblComensales.Caption))
    oCmdEjec.Execute
End Sub

Private Sub lblComensales_Click()
frmDetalle.Comensales = True
frmDetalle.Show vbModal
If frmDetalle.vSelec Then
    Me.lblComensales.Caption = frmDetalle.vDetalle
    If Not VNuevo Then ActualizaDatosComanda
End If
End Sub

Private Sub lvPlatos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Me.lvPlatos.ColumnHeaders(1).Width = 4000
End Sub

Private Sub lvPlatos_DblClick()
Dim oRSdet As ADODB.Recordset
LimpiaParametros oCmdEjec
  oCmdEjec.CommandText = "SpDetalleProducto"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codart", adBigInt, adParamInput, , Me.lvPlatos.SelectedItem.Tag)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TotPa", adInteger, adParamInput, , Me.lvPlatos.SelectedItem.SubItems(3))
   Set oRSdet = oCmdEjec.Execute
    
    If Not oRSdet.EOF Then
    frmDetCombo.lvListado.ColumnHeaders.Add , , "Producto", 4500
    frmDetCombo.lvListado.ColumnHeaders.Add , , "Prom"
    frmDetCombo.lvListado.ColumnHeaders.Add , , "Total"
    frmDetCombo.lvListado.Gridlines = True
    frmDetCombo.lvListado.FullRowSelect = True
    frmDetCombo.lvListado.View = lvwReport
    frmDetCombo.lvListado.LabelEdit = lvwManual
    frmDetCombo.Caption = "COMPOSICIÓN DEL PRODUCTO: " & Me.lvPlatos.SelectedItem.Text
        Do While Not oRSdet.EOF
    
    
            With frmDetCombo.lvListado.ListItems.Add(, , Trim(oRSdet!Prod))
            .SubItems(1) = Format(oRSdet!prom, "##.#0")
            .SubItems(2) = Format(oRSdet!Total, "##.#0")
            End With
            oRSdet.MoveNext
        Loop
        frmDetCombo.Show vbModal
End If

End Sub

Private Sub lvPlatos_ItemClick(ByVal Item As MSComctlLib.ListItem)
'Me.lblSerie.Tag = Me.lvPlatos
End Sub

Private Sub lvPlatos_KeyPress(KeyAscii As Integer)
'Dim i As Integer
'        For i = 1 To Me.lvPlatos.ListItems.count
'
'
'
'        If Me.lvPlatos.ListItems(i).Selected Then
'            LimpiaParametros oCmdEjec
'            oCmdEjec.CommandText = "SpActualizarDetallePlato"
'            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
'            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adChar, adParamInput, 3, Me.lblSerie.Caption)
'            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adDouble, adParamInput, , CDbl(Me.lblNumero.Caption))
'            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numsec", adInteger, adParamInput, , CInt(Me.lvPlatos.ListItems(i).SubItems(6)))
'            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@det", adVarChar, adParamInput, 50, Chr(KeyAscii))
'            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ESDETALLE", adBoolean, adParamInput, , 1)
'            oCmdEjec.Execute
'
'            Me.lvPlatos.ListItems(i).SubItems(1) = Chr(KeyAscii)
'        End If
'        Next
End Sub

Public Sub AgregarDesdeBuscador(xIDproducto As Double, _
                                xProducto As String, _
                                xPrecio As Double)

    Dim c As Integer

    For c = 1 To Me.lvPlatos.ListItems.count
        Me.lvPlatos.ListItems(c).Selected = False
    Next

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

    If vmin Then
        MsgBox "Los siguientes insumos del Plato estan el el Minimo permitido" & vbCrLf & vstrmin, vbInformation, NombreProyecto
    End If

    If vcero Then
        ' MsgBox "Algunos insumos del Plato no estan disponibles" & vbCrLf & vstrcero, vbCritical, NombreProyecto
    End If

    'If vcero Then Exit Sub
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
       
            If AgregaPlato(xIDproducto, 1, FormatNumber(xPrecio, 2), xPrecio, "", "", 0, Me.lblCliente.Caption, IIf(Len(Trim(Me.lblComensales.Caption)) = 0, 0, Me.lblComensales.Caption)) Then
        
                With Me.lvPlatos.ListItems.Add(, , xProducto, Me.ilComanda.ListImages.Item(1).key, Me.ilComanda.ListImages.Item(1).key)
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
                    'ItemP.Checked = True
                    'oRsPlatos.Filter = ""
                    'oRsPlatos.MoveFirst
                  
                End With

                vEstado = "O"
            End If
        End If

    Else

        Dim DD As Integer

        'oRsPlatos.Filter = "Codigo = '" & Me.cmdPlato(Index).Tag & "'"

        If AgregaPlato(xIDproducto, 1, FormatNumber(xPrecio, 2), xPrecio, "", Me.lblSerie.Caption, Me.lblNumero.Caption, Me.lblCliente.Caption, IIf(Len(Trim(Me.lblComensales.Caption)) = 0, 0, frmComanda2.lblComensales.Caption), DD) Then
    
            With Me.lvPlatos.ListItems.Add(, , Me.cmdPlato(index).Caption, Me.ilComanda.ListImages.Item(1).key, Me.ilComanda.ListImages.Item(1).key)
                .Tag = xIDproducto
                .Checked = True
                .SubItems(3) = FormatNumber(1, 2)
                'obteniendo precio
                'oRsPlatos.Filter = "Codigo = '" & Me.cmdPlato(Index).Tag & "'"

                If Not oRsPlatos.EOF Then: .SubItems(4) = FormatNumber(xPrecio, 2)
                .SubItems(5) = FormatNumber(val(.SubItems(3)) * val(.SubItems(4)), 2)
                .SubItems(6) = DD
                .SubItems(7) = 0   'linea nueva
                .SubItems(8) = vMaxFac
                .SubItems(9) = 0
            End With

            'oRsPlatos.Filter = ""
            'oRsPlatos.MoveFirst
    
        End If
    End If

    If Me.lvPlatos.ListItems.count <> 0 Then
        'Me.lvPlatos.SelectedItem = Nothing
        Me.lblTot.Caption = FormatCurrency(sumatoria, 2)
        Me.lblItems.Caption = "Items: " & Me.lvPlatos.ListItems.count
        Me.lvPlatos.ListItems(Me.lvPlatos.ListItems.count).Selected = True
    End If

    'Dim I As Integer
    'prueba

    CargarComanda LK_CODCIA, vMesa
    'aqui

    'Dim C As Integer
    For c = 1 To Me.lvPlatos.ListItems.count
        'If Me.lvPlatos.ListItems(c).Checked Then
        Me.lvPlatos.ListItems(c).Selected = False
        'End If
    Next

    Me.lvPlatos.ListItems(Me.lvPlatos.ListItems.count).Selected = True
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
            If AgregaPlato(oRSbolsa!IDE, oRSbolsa!cant, oRSbolsa!PRE, (oRSbolsa!cant * oRSbolsa!PRE), "", cSerie, cNumero, cCliente, IIf(Len(Trim(Me.lblComensales.Caption)) = 0, 0, Me.lblComensales.Caption), DD2) Then

                With Me.lvPlatos.ListItems.Add(, , oRSbolsa!PRD, Me.ilComanda.ListImages.Item(1).key, Me.ilComanda.ListImages.Item(1).key)
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
