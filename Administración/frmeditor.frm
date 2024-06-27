VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form frmeditor 
   BorderStyle     =   0  'None
   Caption         =   "Editor"
   ClientHeight    =   5970
   ClientLeft      =   0
   ClientTop       =   1005
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   6.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmeditor.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5970
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "&Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
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
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
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
      Left            =   3120
      TabIndex        =   1
      Top             =   5520
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog DIALOGO 
      Left            =   240
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin RichTextLib.RichTextBox DOCU 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   9551
      _Version        =   327680
      BackColor       =   -2147483634
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmeditor.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmeditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VAR

Private Sub Command1_Click()
Static t
Dim inc_page, fin_page, num_copias, i
Dim text As String
Dim a As Integer
Dim cuenta As Integer
Dim suma As Integer
dialogo.CancelError = True
On Error GoTo erro_imp
dialogo.ShowPrinter
inc_page = dialogo.FromPage
fin_page = dialogo.ToPage
num_copias = dialogo.Copies
Dim VAR As Currency

'Printer.KillDoc
Printer.PaperSize = 5
Printer.FontSize = 20 '12
Printer.FontName = "Arial"
Printer.FontBold = False
'Printer.Print " "
Printer.FontSize = 10 '9 '7
Printer.FontName = "Draft 17cpi"
Printer.FontBold = False
'Printer.ScaleHeight = 25000
'Printer.Orientation = 1
If WR_PAG = 999 Then
   Printer.Print DOCU.text
   Printer.EndDoc
   Exit Sub
End If
LEER_TRA_LLAVE
If tra_llave.EOF Then
   Screen.MousePointer = 0
   Exit Sub
End If

For i = 1 To num_copias
 If PUB_CODTRA = 2406 Or PUB_CODTRA = 300 Or PUB_CODTRA = 2401 Then
   Printer.Print DOCU.text
   Printer.EndDoc
 Else
   LKCHEK = False
   If Trim(tra_llave(3)) = "ESTADOGP" Then
      FRMIMP.POWER_REPORT 77
   ElseIf Trim(tra_llave(3)) = "ESTADOGP_2" Then
      FRMIMP.POWER_REPORT 78
   ElseIf WR_PAG = 999 Then
     LKCHEK = False
     FrmImprime.IMPRI_CAMPOS
   End If
 End If
   
   
Next
Exit Sub
erro_imp:
    'Cancelada la impresion
    MsgBox Err.Description
    
    
Exit Sub
End Sub

Private Sub Command2_Click()
Unload frmeditor
End Sub


Private Sub Command3_Click()
dialogo.CancelError = True
On Error GoTo ERR_IMP
dialogo.FLAGS = &H3& Or &H100&
dialogo.ShowFont
frmeditor.DOCU.Font.Name = dialogo.FontName
frmeditor.DOCU.Font.Size = dialogo.FontSize
frmeditor.DOCU.Font.Bold = dialogo.FontBold
frmeditor.DOCU.Font.Italic = dialogo.FontItalic
frmeditor.DOCU.Font.Underline = dialogo.FontUnderline
frmeditor.DOCU.Font.Strikethru = dialogo.FontStrikethru

Exit Sub
ERR_IMP:
 Exit Sub



End Sub

Private Sub Form_Activate()
On Error GoTo SALE

Dim RUTA As String
Screen.MousePointer = 0
If WR_PAG = 999 Then
 RUTA = PUB_RUTA_REPORTE & Trim(FrmImprime.LIST_TABLAS.text) & ".txt"
Else
  RUTA = PUB_RUTA_REPORTE & Trim(LK_CODUSU) & ".txt"
End If
DOCU.Font.Name = "Terminal"
DOCU.Font.Bold = False
DOCU.Font.Size = 8
DOCU.filename = RUTA

SALE:
If Err.Number = 75 Then
   MsgBox "Ruta del Archivo no se encontro ..!!!", 48, Pub_Titulo
   Exit Sub
End If

End Sub

