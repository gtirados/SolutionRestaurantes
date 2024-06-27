VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form FrmAyuda 
   Caption         =   "Ayuda de Transacción"
   ClientHeight    =   5535
   ClientLeft      =   2145
   ClientTop       =   900
   ClientWidth     =   5355
   Icon            =   "FrmAyuda.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5535
   ScaleWidth      =   5355
   Begin VB.CommandButton CmdEscapar 
      Caption         =   "&Escapar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   5040
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog dialogo 
      Left            =   6000
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin RichTextLib.RichTextBox Ayudas 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6376
      _Version        =   327680
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"FrmAyuda.frx":0442
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Diogenes Perez Villalobos S.A."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   345
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Grupo Empresarial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1965
   End
End
Attribute VB_Name = "FrmAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ban As Integer
Private Sub Ayudas_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Unload FrmAyuda
End If
End Sub


Private Sub CmdEscapar_Click()
Unload FrmAyuda
End Sub


Private Sub CmdImprimir_Click()
Dim inc_page, fin_page, num_copias, i
dialogo.CancelError = True
On Error GoTo erro_imp
dialogo.ShowPrinter
inc_page = dialogo.FromPage
fin_page = dialogo.ToPage
num_copias = dialogo.Copies
For i = 1 To num_copias
 Printer.Print Ayudas.text
 Printer.EndDoc
Next
Exit Sub
erro_imp:
    'Cancelada la impresion
Exit Sub

End Sub

Private Sub Form_Load()
Exit Sub
If PUB_CODTRA = 2401 Then
    Ayudas.filename = PUB_RUTA_REPORTE & "h2401.txt"
    Exit Sub
End If
If PUB_CODTRA = 2710 Then
    Ayudas.filename = PUB_RUTA_REPORTE & "h2710.rtf"
    Exit Sub
End If
If PUB_CODTRA = 2403 Then
    Ayudas.filename = PUB_RUTA_REPORTE & "h2403.rtf"
    Exit Sub
End If

'** Pantalla Principal de Transacciones **'
Ayudas.filename = PUB_RUTA_REPORTE & "hprinci.txt"


End Sub




