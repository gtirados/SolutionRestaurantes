VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmCaa 
   BackColor       =   &H00FAEFDA&
   Caption         =   "Estado de Cuenta"
   ClientHeight    =   6255
   ClientLeft      =   210
   ClientTop       =   885
   ClientWidth     =   8970
   ControlBox      =   0   'False
   ForeColor       =   &H00808000&
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6255
   ScaleWidth      =   8970
   WindowState     =   2  'Maximized
   Begin VB.Frame Consulta1 
      BackColor       =   &H00FAEFDA&
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11700
      Begin VB.CommandButton cmdchequeo 
         Caption         =   "Cheque de Saldos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   19
         Top             =   680
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Clientes"
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
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox i_fecha 
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
         Left            =   7800
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox i_codcli 
         Height          =   315
         Left            =   1440
         MaxLength       =   8
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Proveedores"
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
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2880
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label i_limcre 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FAEFDA&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2040
         TabIndex        =   18
         Top             =   600
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label l1 
         BackColor       =   &H00FAEFDA&
         BackStyle       =   0  'Transparent
         Caption         =   "Limite:"
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
         Index           =   1
         Left            =   1440
         TabIndex        =   17
         Top             =   600
         Width           =   855
      End
      Begin VB.Label i_nomCLI 
         BackColor       =   &H00FAEFDA&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   555
         Left            =   2520
         TabIndex        =   12
         Top             =   240
         Width           =   3480
         WordWrap        =   -1  'True
      End
      Begin VB.Label l2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A partir de la Fecha:"
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
         Left            =   7800
         TabIndex        =   5
         Top             =   120
         Width           =   1680
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FAEFDA&
      Height          =   855
      Left            =   0
      TabIndex        =   21
      Top             =   960
      Width           =   11685
      Begin VB.CheckBox periodo 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Acumulado"
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
         Left            =   4335
         TabIndex        =   29
         Top             =   330
         Width           =   1455
      End
      Begin VB.Frame FC 
         BackColor       =   &H00FAEFDA&
         Height          =   615
         Left            =   6120
         TabIndex        =   25
         Top             =   150
         Width           =   3375
         Begin VB.OptionButton OPCONSUL2 
            BackColor       =   &H00FAEFDA&
            Caption         =   "Todas"
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   2
            Left            =   2385
            TabIndex        =   28
            Top             =   240
            Width           =   840
         End
         Begin VB.OptionButton OPCONSUL2 
            BackColor       =   &H00FAEFDA&
            Caption         =   "Canceladas"
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   27
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton OPCONSUL2 
            BackColor       =   &H00FAEFDA&
            Caption         =   "Activas"
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.CommandButton cmddocu 
         Caption         =   "Ver &Documentos"
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
         Height          =   450
         Left            =   9780
         TabIndex        =   23
         Top             =   255
         Width           =   1710
      End
      Begin VB.OptionButton OPCONSUL 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Por Documento "
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
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   24
         Top             =   330
         Value           =   -1  'True
         Width           =   1680
      End
      Begin VB.CheckBox chesub 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Separar por Subtotales"
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
         Left            =   1860
         TabIndex        =   22
         Top             =   330
         Width           =   2400
      End
   End
   Begin ComctlLib.ListView LV_CLI 
      Height          =   855
      Left            =   6900
      TabIndex        =   16
      Tag             =   "0"
      Top             =   7680
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1508
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Estado de Cuenta Grafico en Excel "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   315
      TabIndex        =   15
      Top             =   6405
      Visible         =   0   'False
      Width           =   3060
   End
   Begin MSFlexGridLib.MSFlexGrid GridK 
      Height          =   4170
      Left            =   135
      TabIndex        =   9
      Top             =   2115
      Visible         =   0   'False
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   7355
      _Version        =   393216
      Rows            =   3
      BackColor       =   16777215
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorSel    =   9128212
      GridColorFixed  =   15386515
      GridLinesFixed  =   1
      AllowUserResizing=   3
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
   Begin MSFlexGridLib.MSFlexGrid GRIDG 
      Height          =   4170
      Left            =   120
      TabIndex        =   13
      Top             =   2115
      Visible         =   0   'False
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   7355
      _Version        =   393216
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      GridColorFixed  =   15386515
      GridLinesFixed  =   1
      MergeCells      =   1
      AllowUserResizing=   3
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
   Begin VB.CommandButton CmdImprimir 
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
      Height          =   375
      Left            =   7470
      TabIndex        =   6
      Top             =   6405
      Width           =   1455
   End
   Begin VB.CommandButton SALIR 
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
      Height          =   375
      Left            =   9510
      TabIndex        =   2
      Top             =   6405
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox DOCUM2 
      Height          =   255
      Left            =   5595
      TabIndex        =   8
      Top             =   7845
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"FrmCaa.frx":0000
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
   Begin RichTextLib.RichTextBox DOCUM 
      Height          =   255
      Left            =   8520
      TabIndex        =   7
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"FrmCaa.frx":0076
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
   Begin MSComDlg.CommonDialog dialogo 
      Left            =   6375
      Top             =   7785
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblbarraos 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Solution - Gestión Contable"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   0
      TabIndex        =   30
      Top             =   6900
      Width           =   11955
   End
   Begin VB.Label momento 
      Alignment       =   2  'Center
      Caption         =   "Un Momento . . ."
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
      Left            =   2520
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Nombre 
      AutoSize        =   -1  'True
      Caption         =   "        "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2040
      TabIndex        =   3
      Top             =   960
      Width           =   960
   End
End
Attribute VB_Name = "FrmCaa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim xl As Excel.Application
Dim PSCAA_menor As rdoQuery
Dim caa_menor As rdoResultset
Dim PSCLI_menor As rdoQuery
Dim PSFFF_menor As rdoQuery

Dim cli_menor As rdoResultset
Dim FFF_menor As rdoResultset

Dim pu_saldo As Currency
Dim pu_numdoc
Dim pu_fecha As Date
Dim llave1
Dim pu_flag As Integer
Dim pu_titulo As String
Dim pu_direc As String
Dim pu_numero As String * 10
Dim pu_zona As Integer
Dim pu_subzona As Integer
Dim pu_ruc As String * 12
Dim wcar_mayor As rdoResultset
Dim wPSCAR_MAYOR As rdoQuery
Dim wcar_llave As rdoResultset
Dim wPSCAR_LLAVE As rdoQuery

Dim pu_ultimo As Boolean
Dim pu_maximo As Integer
Dim pu_can As Integer
Dim loc_key As Integer

Private Sub AHORA(GridK As MSFlexGrid)

Dim i, J, JJ
Dim wranF, wran1, wran2
Dim LETRAS(24) As String * 1
'If CmdProcesa.Enabled <> True Then
' MsgBox "Nuestre una Consulta.", 48, Pub_Titulo
'  Exit Sub
'End If
Dim xl As Object
'On Error GoTo FINTODO
Screen.MousePointer = 11
GoSub LETRAS
GoSub WEXCEL
pub_cadena = ""

xl.Cells(4, 1) = "CLIENTE : " & Trim(i_nomCLI.Caption)
xl.Cells(5, 1) = "LIMITE     : " & cli_llave!cli_limcre
xl.Cells(6, 1) = "FECHA APROB: " & cli_llave!CLI_FECHA_APROB
xl.Cells(3, 1) = "'" & Format(LK_FECHA_DIA, "dd/mm/yyyy")
xl.Cells(1, 1) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))

JJ = 0
For i = 0 To GridK.Rows - 1
  For J = 0 To 11
'     If GridK.ColWidth(J) < 100 And JJ = 0 Then
'        wranF = LETRAS(J) & ":" & LETRAS(J)
'        xl.Columns(wranF).ColumnWidth = 0
'       GoTo MAS
'     End If
     xl.Cells(i + 7, J + 1) = "'" & GridK.TextMatrix(i, J)
mas:
  Next J
  JJ = 1
Next i



  'For J = 0 To 11
  '   wranF = LETRAS(J) & ":" & LETRAS(J)
  '   If GridK.ColWidth(J) < 100 Then
  '      xl.Columns(wranF).ColumnWidth = 0
  '   Else
  '      xl.Range(wranF).Select
  '      xl.Columns(wranF).ColumnWidth = GridK.ColWidth(J) / 80
  '   End If
  'Next J


GoSub LETRAS

wranF = "A8:" & "L8"
xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 3

If Left(xl.Cells(7, 11), 4) = "Mont" Then
   wranF = "$A$1:$J$" & i + 7
   xl.Range(wranF).Select
   xl.ActiveSheet.PageSetup.PrintArea = wranF
Else
  xl.Application.Visible = True
   wranF = "$A$1:$L$" & i + 7
   xl.Range(wranF).Select
   xl.ActiveSheet.PageSetup.PrintArea = wranF
'   xl.ActiveWindow.SelectedSheets.PrintOut Copies:=1
End If

xl.Cells(1, 1) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
xl.Cells(2, 1) = "CONSULTA DE CTAS. CTES. "
xl.Cells(3, 1) = "'" & Format(LK_FECHA_DIA, "dd/mm/yyyy")
xl.DisplayAlerts = False
xl.Worksheets(1).Protect PUB_CLAVE
xl.Application.Visible = True
Set xl = Nothing
Screen.MousePointer = 0
Exit Sub

WEXCEL:
  Dim wsfile1
'  lblproceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  'PUB_CLAVE = "IMPOSIBLE"
  xl.Workbooks.Open CONS_ADMIN & "CONTABILIDAD\CTACTE.xls", 0, True, 4
Return

LETRAS:

LETRAS(0) = "A"
LETRAS(1) = "B"
LETRAS(2) = "C"
LETRAS(3) = "D"
LETRAS(4) = "E"
LETRAS(5) = "F"
LETRAS(6) = "G"
LETRAS(7) = "H"
LETRAS(8) = "I"
LETRAS(9) = "J"
LETRAS(10) = "K"
LETRAS(11) = "L"
LETRAS(12) = "M"
LETRAS(13) = "N"
LETRAS(14) = "O"
LETRAS(15) = "P"
LETRAS(16) = "Q"
LETRAS(17) = "R"
LETRAS(18) = "S"
LETRAS(19) = "T"
LETRAS(20) = "U"
LETRAS(21) = "V"
LETRAS(22) = "W"
LETRAS(23) = "X"
Return

FINTODO:
 'MsgBox ERR"Reintente Nuevamente ..", 48, Pub_Titulo
 MsgBox "Reintente Nuevamente ..", 48, Pub_Titulo






End Sub



Private Sub cmddocu_Click()
If Val(i_codcli.Text) <= 0 Then
 Exit Sub
End If
If GRIDG.Visible And GRIDG.Rows > 1 Then
 GRIDG.Visible = False
 GridK.Visible = True
 GridK.SetFocus
 Exit Sub
End If

If Trim(i_fecha.Text) <> "" Then
    If Not IsDate(i_fecha.Text) Then
       MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
       Exit Sub
    End If
    pu_fecha = i_fecha.Text
Else
  pu_fecha = #1/1/1900#

End If
'GridK.Visible = True
GRIDG.Visible = False
pu_flag = 0
Screen.MousePointer = 11
GridK.Visible = False
DoEvents
Momento.Visible = True
DoEvents
If OPCONSUL2(0).Value Then
 pu_titulo = "  ESTADO  DE  CUENTA  ACTIVOS   "
 DOCUMENTO 0
ElseIf OPCONSUL2(1).Value Then
 pu_titulo = "  ESTADO  DE  CUENTA  CANCELADOS "
 DOCUMENTO 1
ElseIf OPCONSUL2(2).Value Then
 DOCUMENTO 2
End If
Momento.Visible = False
DoEvents

If pu_flag = 0 Then
 GridK.Visible = True
 GRIDG.Visible = False
 GridK.Row = 1
 GridK.Col = 2
 GridK.SetFocus
Else
If PUB_CP = "P" Then
  MsgBox "NO tiene Cuenta x Cobrar Pendiente...", 48, Pub_Titulo
Else
  MsgBox "NO tiene Cuenta x Pagar Pendiente...", 48, Pub_Titulo
End If
 GRIDG.Visible = False
 GridK.Visible = False
 cmddocu.SetFocus
End If
Screen.MousePointer = 0

End Sub

Private Sub cmddocu_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  FrmCaa.GridK.Visible = False
  FrmCaa.i_codcli.SetFocus
End If

End Sub

Private Sub CmdImprimir_Click()
If GridK.Visible = True Then AHORA GridK
If GRIDG.Visible = True Then AHORA GRIDG

End Sub

Private Sub cmdlimcre_Click()
Dim PSLIM As rdoQuery
Dim ps_limcre As rdoResultset
Dim Mensaje, Título, valorpred, mifecha
If Val(i_codcli.Text) = 0 Then i_codcli.SetFocus: Exit Sub
Mensaje = "Ingrese una Fecha de Inicio para la Consulta : "
Título = "Movimientos de Limite de Credito de Clientes"
valorpred = Format(LK_FECHA_DIA, "dd/mm/yyyy")
mifecha = InputBox(Mensaje, Título, valorpred)
If mifecha = "" Then
   Exit Sub
End If
If Not IsDate(mifecha) Then
   MsgBox "Fecha Invalidad , Intente Nuevamente", 48, Pub_Titulo
   Exit Sub
End If

pub_cadena = "SELECT ALL_CODTRA,ALL_FECHA_DIA, ALL_LIMCRE_ACT,ALL_LIMCRE_ANT, ALL_CODUSU ,ALL_IMPORTE , ALL_HORA FROM ALLOG WHERE ALL_FECHA_DIA >= ? and ALL_CODCLIE = ? and ALL_CODCIA = ? and  ALL_CODTRA = 2580 ORDER BY ALL_FECHA_DIA, ALL_NUMOPER"
Set PSLIM = CN.CreateQuery("", pub_cadena)
PSLIM.rdoParameters(0) = LK_FECHA_DIA
PSLIM.rdoParameters(1) = 0
PSLIM.rdoParameters(2) = " "
Set ps_limcre = PSLIM.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
On Error GoTo OJO
PSLIM(0) = mifecha
PSLIM(1) = Val(i_codcli.Text)
PSLIM(2) = LK_CODCIA
ps_limcre.Requery
On Error GoTo 0
If ps_limcre.EOF Then
 MsgBox " No Existe Movimientos de Consulta.", 48, Pub_Titulo
 Exit Sub
End If
Screen.MousePointer = 11
GridK.Visible = False
GRIDG.Visible = False
Momento.Visible = True
DoEvents
FrmCaa.GRIDG.Clear
FrmCaa.GRIDG.Cols = 5
FrmCaa.GRIDG.ColWidth(0) = 1200
FrmCaa.GRIDG.ColWidth(1) = 2500
FrmCaa.GRIDG.ColWidth(2) = 2500
FrmCaa.GRIDG.ColWidth(3) = 1300
FrmCaa.GRIDG.ColWidth(4) = 1300
FrmCaa.GRIDG.TextMatrix(0, 0) = "Fecha"
FrmCaa.GRIDG.TextMatrix(0, 1) = "Limite Efectuado"
FrmCaa.GRIDG.TextMatrix(0, 2) = "Limite Anterior"
FrmCaa.GRIDG.TextMatrix(0, 3) = "Usuario"
FrmCaa.GRIDG.TextMatrix(0, 4) = "Hora"
fila = 0
FrmCaa.GRIDG.Rows = 1
Do Until ps_limcre.EOF
  fila = fila + 1
  FrmCaa.GRIDG.Rows = FrmCaa.GRIDG.Rows + 1
  FrmCaa.GRIDG.TextMatrix(fila, 0) = ps_limcre!ALL_FECHA_DIA
  FrmCaa.GRIDG.TextMatrix(fila, 1) = Nulo_Valor0(ps_limcre!ALL_LIMCRE_ACT)
  FrmCaa.GRIDG.TextMatrix(fila, 2) = Nulo_Valor0(ps_limcre!ALL_LIMCRE_ANT)
  FrmCaa.GRIDG.TextMatrix(fila, 3) = ps_limcre!all_codusu
  FrmCaa.GRIDG.TextMatrix(fila, 4) = Format(ps_limcre!ALL_HORA, "hh:mm:ss AMPM")
  ps_limcre.MoveNext
Loop

Screen.MousePointer = 0
GRIDG.Visible = True
GRIDG.Col = 1
GRIDG.Row = 1
GRIDG.SetFocus
Momento.Visible = False
DoEvents
Exit Sub
OJO:
MsgBox "Intente Nuevamente.", 48, Pub_Titulo
End Sub


Private Sub Command1_Click()
If GRIDG.Visible = False And GridK.Visible = False Then
  Exit Sub
End If
ACU_GRAF
If pu_flag = 1 Then
 MsgBox "NO Existe Movimentos"
End If
End Sub



Private Sub GRIDG_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  FrmCaa.GRIDG.Visible = False
  FrmCaa.i_codcli.SetFocus
End If

If KeyAscii <> 13 Then
  Exit Sub
End If

If OPCONSUL(1).Value = True Or Option1(1).Value Then
  Exit Sub
End If
 Dim WSUM As Currency
 Dim TEXTO
 TEXTO = ""
 Dim pb
 pb = Chr(10) & Chr(13)
 If GRIDG.Col = 2 Then
   If LK_EMP = "PLA" Then
     SQ_OPER = 1
     pu_codcia = LK_CODCIA
     'PUB_CODBAN = wcaA_mayor!caa_codban
     'LEER_CCM_LLAVE
     'If Not ccm_llave.EOF Then MsgBox Trim(ccm_llave!CCM_nombre)
   Else
      Exit Sub
    PSFFF_menor.rdoParameters(0) = LK_CODCIA
    PSFFF_menor.rdoParameters(1) = GRIDG.TextMatrix(GRIDG.Row, 12)
    PSFFF_menor.rdoParameters(2) = GRIDG.TextMatrix(GRIDG.Row, 11)
    PSFFF_menor.rdoParameters(3) = GRIDG.TextMatrix(GRIDG.Row, 13)
    FFF_menor.Requery
    FFF_menor.MoveFirst
    TEXTO = ""
    Do Until FFF_menor.EOF
       PUB_KEY = FFF_menor!far_codart
       pu_codcia = LK_CODCIA
       SQ_OPER = 1
       LEER_ART_LLAVE
       TEXTO = TEXTO + Left(art_LLAVE!ART_NOMBRE, 20) & " : " & FFF_menor!FAR_descri & " : " & Format((FFF_menor!FAR_CANTIDAD / FFF_menor!FAR_equiv), "0.00") + pb
       WSUM = WSUM + FFF_menor!FAR_CANTIDAD / FFF_menor!FAR_equiv
       FFF_menor.MoveNext
    Loop
    MsgBox "Contenido de Celda : " + pb + TEXTO + pb + "Total =   " + Format(WSUM, "0.00"), vbInformation, Pub_Titulo
    End If
    GRIDG.SetFocus
    Exit Sub
End If
 If GRIDG.Col = 8 Then
    PUB_CODVEN = GRIDG.TextMatrix(GRIDG.Row, 8)
    pu_codcia = LK_CODCIA
    SQ_OPER = 1
    LEER_VEN_LLAVE
    If Not ven_llave.EOF Then
    TEXTO = ven_llave!VEM_NOMBRE
    MsgBox "Contenido de Celda : " + pb + TEXTO, vbInformation, Pub_Titulo
    GRIDG.SetFocus
    End If
    Exit Sub
End If
If GRIDG.Col = 10 Then
 MsgBox "Contenido de Celda : " + pb + Format(GRIDG.Text, "hh:mm:ss AMPM"), 48, Pub_Titulo
 GRIDG.SetFocus
 Exit Sub
End If
If GRIDG.Col = 9 Then
 usu.Requery
 If usu.EOF Then
   MsgBox "Error de Usuarios Vuelva a Ingresar al Sistema", vbCritical, Pub_Titulo
   End
 End If
 Do Until usu.EOF
  If Trim(usu!usu_key) = Trim(GRIDG.Text) Then
    MsgBox "Contenido de Celda : " + pb + Trim(usu!USU_NOMBRE), vbInformation, Pub_Titulo
    GRIDG.SetFocus
    Exit Sub
  End If
  usu.MoveNext
 Loop
 GRIDG.SetFocus
 Exit Sub
End If

MsgBox "Contenido de Celda : " + pb + GRIDG.TextMatrix(GRIDG.Row, GRIDG.Col), vbInformation, Pub_Titulo
GRIDG.SetFocus

End Sub

Private Sub GridK_DblClick()
 Dim pub_mensajeText
 Dim pb
 pb = Chr(10) & Chr(13) & Chr(10) & Chr(13)
 MsgBox "Contenido de Celda : " + pb + GridK.TextMatrix(GridK.Row, GridK.Col), vbInformation, Pub_Titulo
End Sub

Private Sub i_codcli_Change()
pu_codclie = Val(i_codcli.Text)
End Sub

Private Sub i_codcli_GotFocus()
GRIDG.Visible = False
GridK.Visible = False
'FrmCaa.i_nomCLI.Caption = ""
FrmCaa.i_limcre.Caption = ""
'Azul i_codcli, i_codcli

End Sub
Private Sub Form_Load()
Dim fech As String
Dim fecha As String
If LK_CODUSU = "ADMIN" Then
    cmdchequeo.Visible = True
End If
pub_cadena = "SELECT MOV_TIPO_CAMBIO, MOV_FLAG_TC, MOV_DH, MOV_FECHA_EMI, MOV_MONEDA, MOV_IMPORTE, MOV_NUMFAC, MOV_SERIE, MOV_SUNAT, MOV_DETALLE, MOV_FBG , MOV_NRO_VOUCHER  FROM MOVICONT WHERE  MOV_CODCIA = ? AND MOV_CP = ? AND MOV_CODCLIE = ?  AND (MOV_NRO_MES >= ? and MOV_NRO_MES <= ? )  ORDER BY MOV_CODCIA, MOV_FECHA_EMI"
Set wPSCAR_MAYOR = CN.CreateQuery("", pub_cadena)
wPSCAR_MAYOR.rdoParameters(0) = " "
wPSCAR_MAYOR.rdoParameters(1) = " "
wPSCAR_MAYOR.rdoParameters(2) = 0
wPSCAR_MAYOR.rdoParameters(3) = 0
wPSCAR_MAYOR.rdoParameters(4) = 0
Set wcar_mayor = wPSCAR_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)

pu_ultimo = False
FrmCaa.GridK.Cols = 16
FrmCaa.GridK.Rows = 100
FrmCaa.GridK.ColWidth(0) = 800
FrmCaa.GridK.ColWidth(1) = 1500
FrmCaa.GridK.ColWidth(2) = 900
FrmCaa.GridK.ColWidth(3) = 900
FrmCaa.GridK.ColWidth(4) = 900
FrmCaa.GridK.ColWidth(5) = 900 '
FrmCaa.GridK.ColWidth(6) = 900
FrmCaa.GridK.ColWidth(7) = 900
FrmCaa.GridK.ColWidth(8) = 700
FrmCaa.GridK.ColWidth(9) = 700
FrmCaa.GridK.ColWidth(10) = 2500
Dim cade As String

cade = "SELECT * FROM CARACU WHERE CAA_CP = ?  AND CAA_CODCLIE = ? AND CAA_CODCIA = ? AND CAA_FECHA >= ? AND CAA_NUMDOC = ? AND CAA_SERDOC = ? AND CAA_TIPDOC = ? ORDER BY CAA_CP, CAA_CODCLIE, CAA_CODCIA, CAA_FECHA, CAA_NUM_OPER"
Set PSCAA_menor = CN.CreateQuery("", cade)
PSCAA_menor.rdoParameters(0) = " "
PSCAA_menor.rdoParameters(1) = 0
PSCAA_menor.rdoParameters(2) = " "
PSCAA_menor.rdoParameters(3) = LK_FECHA_DIA
PSCAA_menor.rdoParameters(4) = 0
PSCAA_menor.rdoParameters(5) = 0
PSCAA_menor.rdoParameters(6) = " "
Set caa_menor = PSCAA_menor.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)

'cade = "SELECT * FROM CLIENTES WHERE Cli_NOMBRE >= ?  ORDER BY CLI_NOMBRE"
'Set PSCLI_menor = CN.CreateQuery("", cade)
'Set cli_menor = PSCLI_menor.OpenResultset(rdOpenKeyset, rdConcurValues)

cade = "SELECT * FROM FACART WHERE FAR_CODCIA= ?  AND FAR_NUMSER = ? AND FAR_FBG = ? AND FAR_NUMFAC = ? ORDER BY FAR_NUMSEC"
Set PSFFF_menor = CN.CreateQuery("", cade)
PSFFF_menor.rdoParameters(0) = " "
PSFFF_menor.rdoParameters(1) = 0
PSFFF_menor.rdoParameters(2) = " "
PSFFF_menor.rdoParameters(3) = 0

Set FFF_menor = PSFFF_menor.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)


FrmCaa.GridK.Col = 1
FrmCaa.GridK.ColAlignment(1) = 0
FrmCaa.GridK.Col = 2
FrmCaa.GridK.ColAlignment(2) = 0
FrmCaa.GridK.Col = 3
FrmCaa.GridK.ColAlignment(3) = 1
FrmCaa.GridK.Col = 4
FrmCaa.GridK.ColAlignment(4) = 2
FrmCaa.GridK.Col = 5
FrmCaa.GridK.ColAlignment(5) = 1
fecha = DateAdd("m", -1, LK_FECHA_DIA)
fech = Format(fecha, "mm")
i_fecha.Text = "01" & "/" & fech & "/" & Right(Format(fecha, "yyyy"), 4)
LV_CLI.Width = 3000
LV_CLI.Height = 3000
LV_CLI.Top = 1000
LV_CLI.Left = 3000
'FrmCaa.Agrupar.Value = 0
' OPCONSUL(0).Value = False
' OPCONSUL(1).Value = True
'OPCONSUL_Click 1
i_codcli.TabIndex = 0
cmddocu.Enabled = True
OPCONSUL2(0).Enabled = True
OPCONSUL2(0).Value = True

OPCONSUL2(1).Enabled = True
OPCONSUL2(2).Enabled = True
'ultimo.Enabled = False
End Sub


Private Sub GRIDK_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  FrmCaa.GridK.Visible = False
  FrmCaa.i_codcli.SetFocus
End If

If KeyAscii = 13 Then
Exit Sub
 'GridK.Visible = False
  If Trim(GridK.TextMatrix(GridK.Row, 2)) = "" Then
  Exit Sub
  End If
wPSCAR_LLAVE.rdoParameters(0) = LK_CODCIA
wPSCAR_LLAVE.rdoParameters(1) = pu_cp
wPSCAR_LLAVE.rdoParameters(2) = Val(i_codcli.Text)
wPSCAR_LLAVE.rdoParameters(3) = Trim(GridK.TextMatrix(GridK.Row, 10))
wPSCAR_LLAVE.rdoParameters(4) = Val(GridK.TextMatrix(GridK.Row, 11))
wPSCAR_LLAVE.rdoParameters(5) = Val(GridK.TextMatrix(GridK.Row, 12))

pu_flag = 0
DETALLE
If pu_flag = 0 Then
   GridK.Visible = False
   GRIDG.Visible = True
   GRIDG.SetFocus
Else
   GridK.SetFocus
 End If
End If
End Sub


Private Sub i_codcli_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If KeyCode = 113 Then
    If Option1(0).Value Then
      Option1(0).Value = False
      Option1(1).Value = True
      Option1_Click (1)
    Else
     Option1(0).Value = True
      Option1(1).Value = False
      Option1_Click (2)
    End If
    Exit Sub
End If

If Not LV_CLI.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And i_codcli.Text = "" Then
  loc_key = 1
  Set LV_CLI.SelectedItem = LV_CLI.ListItems(loc_key)
'  LV_CLI.Visible = False
  LV_CLI.ListItems.Item(loc_key).Selected = True
  LV_CLI.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > LV_CLI.ListItems.Count Then loc_key = LV_CLI.ListItems.Count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > LV_CLI.ListItems.Count Then loc_key = LV_CLI.ListItems.Count
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
  LV_CLI.ListItems.Item(loc_key).Selected = True
  LV_CLI.ListItems.Item(loc_key).EnsureVisible
  i_codcli.Text = Trim(LV_CLI.ListItems.Item(loc_key).Text) & " "
  DoEvents
  i_codcli.SelStart = Len(i_codcli.Text)
  DoEvents
fin:

End Sub

Private Sub i_codcli_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
If Len(i_codcli.Text) = 0 Or IsNumeric(i_codcli.Text) = True Then
   LV_CLI.Visible = False
   Exit Sub
End If
If LV_CLI.Visible = False And KeyCode <> 13 Or Len(i_codcli.Text) = 1 Then
    var = Asc(i_codcli.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    numarchi = 1
    archi = "SELECT CLI_CODCLIE , CLI_CODCIA, CLI_CP, CLI_NOMBRE , CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM  FROM CLIENTES WHERE CLI_CP = '" & pu_cp & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & i_codcli.Text & "' AND  '" & var & "' ORDER BY CLI_NOMBRE"
    PROC_LISVIEW LV_CLI
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If LV_CLI.Visible Then
  Set itmFound = LV_CLI.FindItem(LTrim(i_codcli.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > LV_CLI.ListItems.Count Then
      LV_CLI.ListItems.Item(LV_CLI.ListItems.Count).EnsureVisible
   Else
     LV_CLI.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If


End Sub


Private Sub i_fecha_KeyPress(KeyAscii As Integer)
Dim d As Integer
If KeyAscii <> 13 Then
Exit Sub
End If
 If Trim(i_fecha.Text) <> "" Then
    If Not IsDate(i_fecha) Then
     MsgBox "Fecha es Invalidad  ....!", 48
     Azul i_fecha, i_fecha
     Exit Sub
    End If
 End If
  
End Sub
Private Sub i_codcli_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As ListItem    ' Variable FoundItem.
FrmCaa.i_nomCLI.Caption = ""
If FrmCaa.Option1(0).Value Then
 pu_cp = "C"
Else
 pu_cp = "P"
End If
If KeyAscii = 27 Then
  LV_CLI.Visible = False
  i_codcli.Text = ""
  Exit Sub
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
pu_codclie = Val(i_codcli.Text)
If Len(i_codcli.Text) = 0 Then
   Exit Sub
End If
If pu_codclie <> 0 Then
   If Len(Trim(i_codcli.Text)) = LK_DIG_RUC Then ' LONG DEL RUC
        'pu_cp = Left(CmbCGP.Text, 1)
        PUB_RUC = Trim(i_codcli.Text)
        SQ_OPER = 4
        pu_codcia = LK_CODCIA
        LEER_CLI_LLAVE
        If cli_ruc.EOF Then
           MsgBox "R.U.C. No Existe ", 48, Pub_Titulo
           Exit Sub
        End If
        i_codcli.Text = cli_ruc!CLI_CODCLIE
   End If
   SQ_OPER = 1
   pu_codclie = Val(i_codcli.Text)
   pu_codcia = LK_CODCIA
   LEER_CLI_LLAVE
   If cli_llave.EOF Then
     Azul i_codcli, i_codcli
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     GoTo fin
   Else
      i_nomCLI.Caption = Trim(cli_llave(3)) & Chr(13) & "RUC. " & cli_llave!cli_ruc_esposo
   End If
   cmddocu.SetFocus
   cmddocu_Click
Else
   If loc_key > LV_CLI.ListItems.Count Or loc_key = 0 Then
     Exit Sub
   End If
   valor = UCase(LV_CLI.ListItems.Item(loc_key).Text)
   If Trim(UCase(i_codcli.Text)) = Left(valor, Len(Trim(i_codcli.Text))) Then
   Else
      Exit Sub
   End If
   i_codcli.Text = Trim(LV_CLI.ListItems.Item(loc_key).SubItems(1))
   pu_codclie = Val(i_codcli.Text)
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_CLI_LLAVE
   cmddocu.SetFocus
   cmddocu_Click
End If
LV_CLI.Visible = False
       FrmCaa.i_nomCLI.Caption = Trim(cli_llave(3)) & Chr(13) & "RUC. " & cli_llave!cli_ruc_esposo
       pu_direc = Trim(cli_llave(10))
       pu_numero = Trim(cli_llave(11))
       pu_zona = cli_llave(12)
       pu_subzona = cli_llave(13)
       pu_ruc = cli_llave!cli_ruc_esposo
       

fin:
End Sub


Private Sub LV_CLI_DblClick()
 'i_codcli.SetFocus
 loc_key = LV_CLI.SelectedItem.Index
 i_codcli.Text = Trim(LV_CLI.ListItems.Item(loc_key).Text) & " "
 'i_nomCLI.Caption = LV_CLI.ListItems(loc_key)
 i_codcli_KeyPress 13
 'i_codcli.SetFocus
 
End Sub

Private Sub LV_CLI_GotFocus()
If loc_key <> 0 Then
 Set LV_CLI.SelectedItem = LV_CLI.ListItems(loc_key)
 LV_CLI.ListItems.Item(loc_key).Selected = True
 LV_CLI.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub LV_CLI_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = LV_CLI.SelectedItem.Index
 i_codcli.Text = Trim(LV_CLI.ListItems.Item(loc_key).Text) & " "
End If

End Sub

Private Sub LV_CLI_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 LV_CLI.Visible = False
 i_codcli.Text = ""
 i_codcli.SetFocus
 Exit Sub
End If
If KeyAscii <> 13 Then
 Exit Sub
End If
LV_CLI_DblClick

End Sub

Private Sub LV_CLI_LostFocus()
LV_CLI.Visible = False
End Sub

Private Sub OPCONSUL_Click(Index As Integer)
GridK.Visible = False
GRIDG.Visible = False
If Index = 0 Then
    'OPCONSUL(1).Enabled = True
    'OPCONSUL2(0).Enabled = False
    'OPCONSUL2(1).Enabled = False

ElseIf Index = 1 Then
    'OPCONSUL2(1).Enabled = True
    'OPCONSUL2(0).Value = True
'    cmddocu.Enabled = True
'    cmddocu_Click
End If

End Sub

Private Sub OPCONSUL2_Click(Index As Integer)
GridK.Visible = False
GRIDG.Visible = False

End Sub

Private Sub Option1_Click(Index As Integer)
GRIDG.Visible = False
GridK.Visible = False
FrmCaa.i_codcli.Text = ""
FrmCaa.i_fecha.Text = ""
FrmCaa.i_nomCLI.Caption = ""
FrmCaa.i_codcli.SetFocus
End Sub

Private Sub salir_Click()
 Unload FrmCaa
End Sub

Public Sub cabe()
FrmCaa.GridK.TextMatrix(0, 0) = "Fecha"
FrmCaa.GridK.TextMatrix(0, 1) = "Concepto"
FrmCaa.GridK.TextMatrix(0, 2) = "Ingreso"
FrmCaa.GridK.TextMatrix(0, 3) = "Salida"
FrmCaa.GridK.TextMatrix(0, 4) = "Saldo"
FrmCaa.GridK.TextMatrix(0, 5) = "Ingreso"
FrmCaa.GridK.TextMatrix(0, 6) = "Salida"
FrmCaa.GridK.TextMatrix(0, 7) = "Saldo"
FrmCaa.GridK.TextMatrix(0, 8) = "Cost.Prom."
FrmCaa.GridK.TextMatrix(0, 9) = "Precio"
FrmCaa.GridK.TextMatrix(0, 10) = "Tr."
FrmCaa.GridK.TextMatrix(0, 11) = "Nombre"
FrmCaa.GridK.TextMatrix(0, 12) = "Vendedor"
FrmCaa.GridK.TextMatrix(0, 13) = "Dias"
FrmCaa.GridK.TextMatrix(0, 14) = "N.frmcaa"
FrmCaa.GridK.TextMatrix(0, 15) = "Cia."

End Sub

Public Sub DOCUMENTO(wdocu As Integer)
Dim vdocum
Dim cIngreso As Currency
Dim cSalida As Currency
Dim cSaldo As Currency
Dim fila1 As Integer
Dim PRECIO As Currency
Dim cuenta As String * 6
Dim articulo As String * 6
Dim FINAL As String * 1
Dim Num_Fac As String * 5
Dim Num_Ser As String * 5
Dim WW_CODART As Long
Dim wFAR_CODART
Dim fila As Integer
Dim WS_SALDO As Currency
Dim WS_SALDO2  As Currency
Dim Band As String * 1
Dim xCODCIA As String
Dim vCONCEPTO
Dim LQ_TIPDOC
Dim LQ_MONEDA_2
Dim SUM_TIPDOC As Currency
Dim LQ_MONEDA As String
Dim SUM_MONEDA As Currency
Dim wc_acuenta As Currency
Dim wc_dh1 As String * 1
Dim wc_dh2 As String * 1
Dim WS_TIPO_CAMBIO As Currency
If pu_cp = "C" Then
 wc_dh1 = "H"
 wc_dh2 = "D"
Else
 wc_dh1 = "D"
 wc_dh2 = "H"
End If
'xCODCIA = "SELECT MOV_FECHA_EMI, MOV_MONEDA, MOV_IMPORTE, MOV_NUMFAC, MOV_SERIE, MOV_SUNAT, MOV_DETALLE, MOV_FBG FROM MOVICONT WHERE  MOV_CODCIA = ? AND MOV_CP = ? AND MOV_CODCLIE = ? AND MOV_FBG_C = ? AND MOV_SERIE_C = ?  AND MOV_NUMFAC_C = ? AND MOV_NRO_MES = " & LK_NRO_MES & " ORDER BY MOV_FECHA_EMI"
'Set wPSCAR_LLAVE = CN.CreateQuery("", xCODCIA)
'wPSCAR_LLAVE.rdoParameters(0) = " "
'wPSCAR_LLAVE.rdoParameters(1) = " "
'wPSCAR_LLAVE.rdoParameters(2) = 0
'wPSCAR_LLAVE.rdoParameters(3) = 0
'wPSCAR_LLAVE.rdoParameters(4) = 0
'wPSCAR_LLAVE.rdoParameters(5) = 0
'Set wcar_llave = wPSCAR_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)


'xCODCIA = "SELECT * FROM MOVICONT WHERE  MOV_CODCIA = ? AND MOV_CP = ? AND MOV_CODCLIE = ? AND MOV_MARCA = 'X'  AND MOV_NRO_MES = " & LK_NRO_MES & " ORDER BY MOV_CODCIA, MOV_CP, MOV_FECHA"
Screen.MousePointer = 11
SQ_OPER = 2
wPSCAR_MAYOR(0) = LK_CODCIA
wPSCAR_MAYOR(1) = pu_cp
wPSCAR_MAYOR(2) = Val(i_codcli.Text)
If periodo.Value = 0 Then
    wPSCAR_MAYOR(3) = LK_NRO_MES
    wPSCAR_MAYOR(4) = LK_NRO_MES
Else
    wPSCAR_MAYOR(3) = 1
    wPSCAR_MAYOR(4) = 12
End If

wcar_mayor.Requery
' DETALLE DE CUENTA
'wPSCAR_LLAVE.rdoParameters(0) = LK_CODCIA
'wPSCAR_LLAVE.rdoParameters(1) = pu_cp
'wPSCAR_LLAVE.rdoParameters(2) = Val(i_codcli.Text)

If wcar_mayor.EOF Then
'   pu_flag = 1
'   GoTo SIGUE
End If

FrmCaa.GridK.Font.Size = 8
FrmCaa.GridK.Clear
FrmCaa.GridK.Cols = 15
FrmCaa.GridK.ColAlignment(4) = 1
FrmCaa.GridK.ColWidth(0) = 950
FrmCaa.GridK.ColWidth(1) = 0 '     950
FrmCaa.GridK.ColWidth(2) = 1500
FrmCaa.GridK.ColWidth(3) = 2200
FrmCaa.GridK.ColWidth(4) = 2100

FrmCaa.GridK.ColWidth(5) = 0
FrmCaa.GridK.ColWidth(6) = 500
FrmCaa.GridK.ColWidth(7) = 1
FrmCaa.GridK.ColWidth(8) = 1
FrmCaa.GridK.ColWidth(9) = 1400

FrmCaa.GridK.ColWidth(10) = 1400

FrmCaa.GridK.ColWidth(11) = 1400
FrmCaa.GridK.ColWidth(12) = 0
FrmCaa.GridK.ColWidth(13) = 0
FrmCaa.GridK.ColWidth(14) = 0


FrmCaa.GridK.TextMatrix(0, 0) = "Fec.Proc."
FrmCaa.GridK.TextMatrix(0, 1) = "Fec.Emis."
FrmCaa.GridK.TextMatrix(0, 2) = "Tip.Doc."
FrmCaa.GridK.TextMatrix(0, 3) = "Documento"
FrmCaa.GridK.TextMatrix(0, 4) = "Glosa"

FrmCaa.GridK.TextMatrix(0, 5) = "Fec.Vcto"
FrmCaa.GridK.CellFontBold = True
FrmCaa.GridK.CellForeColor = QBColor(12)
FrmCaa.GridK.TextMatrix(0, 9) = "Debe"
FrmCaa.GridK.TextMatrix(0, 10) = "Haber"
FrmCaa.GridK.Row = 0
FrmCaa.GridK.Col = 11
FrmCaa.GridK.TextMatrix(0, 11) = "Saldo Actual."
FrmCaa.GridK.TextMatrix(0, 12) = " B a n c o. "
FrmCaa.GridK.TextMatrix(0, 13) = "Ved."


fila = 0
FrmCaa.GridK.Rows = 1
cSaldo = 0
SUM_TIPDOC = 0
SUM_MONEDA = 0
wc_acuenta = 0
'LQ_TIPDOC = wcar_mayor!CAR_TIPDOC
'LQ_MONEDA_2 = wcar_mayor!CAR_MONEDA
FrmCaa.GridK.Rows = FrmCaa.GridK.Rows + 1
SQ_OPER = 5
pu_codclie = Val(i_codcli.Text)
pu_codcia = LK_CODCIA
LEER_CLI_LLAVE
cSaldo = 0
GridK.TextMatrix(FrmCaa.GridK.Rows - 1, 4) = "SALDO INICIAL ="
GridK.TextMatrix(FrmCaa.GridK.Rows - 1, 6) = " S/."
If periodo.Value = 0 Then
 JALA_SALDO_CLI pu_codclie, pu_cp, 3
Else
 JALA_SALDO_CLI pu_codclie, pu_cp, 10
End If
'GridK.TextMatrix(fila, 9) = Format(Val(cls_llaver!MOV_IMPORTE) - Val(wc_acuenta), "0.00")
GridK.TextMatrix(FrmCaa.GridK.Rows - 1, 9) = Format(Val(PUB_IMPORTE_HAB), "0.00")
GridK.TextMatrix(FrmCaa.GridK.Rows - 1, 10) = Format(Val(PUB_IMPORTE_DEB), "0.00")

cSaldo = Val(GridK.TextMatrix(FrmCaa.GridK.Rows - 1, 9)) - Val(GridK.TextMatrix(FrmCaa.GridK.Rows - 1, 10))
If pu_cp = "C" Then
GridK.TextMatrix(FrmCaa.GridK.Rows - 1, 11) = Format(Val(PUB_IMPORTE_DEB) - Val(PUB_IMPORTE_HAB), "0.00")
Else
GridK.TextMatrix(FrmCaa.GridK.Rows - 1, 11) = Format(Val(PUB_IMPORTE_HAB) - Val(PUB_IMPORTE_DEB), "0.00")
End If
fila = fila + 1
WS_TIPO_CAMBIO = 1
Do Until wcar_mayor.EOF
'  If wcar_mayor!MOV_NRO_VOUCHER = 462 Then Stop
  fila = fila + 1
  FrmCaa.GridK.Rows = FrmCaa.GridK.Rows + 1
  If wcar_mayor!MOV_MONEDA = "D" Then
    LQ_MONEDA = "$ "
  Else
    LQ_MONEDA = "S/."
  End If
  GridK.Row = fila
  GridK.TextMatrix(fila, 0) = Format(wcar_mayor!MOV_fecha_EMI, "dd/mm/yy")
  GridK.TextMatrix(fila, 2) = wcar_mayor!MOV_SUNAT
  vCONCEPTO = Trim(wcar_mayor!MOV_DETALLE)
  vdocum = Format(wcar_mayor!MOV_serie, "000") & "-" & wcar_mayor!MOV_numfac & " /Vouc.: " & wcar_mayor!MOV_NRO_VOUCHER
  GridK.TextMatrix(fila, 3) = vdocum
  GridK.TextMatrix(fila, 4) = vCONCEPTO
  WS_TIPO_CAMBIO = 1
  If wcar_mayor!MOV_MONEDA = "D" Then
     SQ_OPER = 1
     PUB_CAL_INI = wcar_mayor!MOV_fecha_EMI
     PUB_CAL_FIN = wcar_mayor!MOV_fecha_EMI
     PUB_CODCIA = LK_CODCIA
     LEER_CAL_LLAVE
     If Nulo_Valor0(cal_llave!cal_tipo_cambio) = 0 Then
       MsgBox "Falta tipo de cambio" & wcar_mayor!MOV_fecha_EMI, 48, Pub_Titulo
       Exit Sub
     End If
     WS_TIPO_CAMBIO = Val(cal_llave!cal_tipo_cambio)
  End If
  If wcar_mayor!MOV_FLAG_TC = "A" Then WS_TIPO_CAMBIO = Val(wcar_mayor!MOV_TIPO_CAMBIO)
    GridK.TextMatrix(fila, 6) = " S/."
    If wcar_mayor!MOV_DH = "H" Then
     GridK.TextMatrix(fila, 9) = Format(Val(wcar_mayor!MOV_IMPORTE) * WS_TIPO_CAMBIO, "0.00")
     cSaldo = cSaldo + Val(GridK.TextMatrix(fila, 9))
    Else
     GridK.TextMatrix(fila, 10) = Format(Val(wcar_mayor!MOV_IMPORTE) * WS_TIPO_CAMBIO, "0.00")
     cSaldo = cSaldo - Val(GridK.TextMatrix(fila, 10))
    End If
    'cSaldo = cSaldo Val(GridK.TextMatrix(FrmCaa.GridK.Rows - 1, 9)) - Val(GridK.TextMatrix(FrmCaa.GridK.Rows - 1, 10))
    'If pu_cp = "C" Then
    GridK.TextMatrix(fila, 11) = Format(cSaldo, "0.00") 'Format(Val(GridK.TextMatrix(fila, 9)) - Val(GridK.TextMatrix(fila, 10)), "0.00")
    'Else
    'GridK.TextMatrix(fila, 11) = Format(Val(GridK.TextMatrix(fila, 10)) - Val(GridK.TextMatrix(fila, 9)), "0.00")
    'End If
  GridK.TextMatrix(fila, 12) = wcar_mayor!MOV_FBG
  GridK.TextMatrix(fila, 13) = wcar_mayor!MOV_serie
  GridK.TextMatrix(fila, 14) = wcar_mayor!MOV_numfac
  SUM_TIPDOC = SUM_TIPDOC + Val(GridK.TextMatrix(fila, 9))
  SUM_MONEDA = SUM_MONEDA + Val(GridK.TextMatrix(fila, 9))
 ' cSaldo = (wcar_mayor!MOV_IMPORTE * WS_TIPO_CAMBIO) + cSaldo
valeloop:
  wcar_mayor.MoveNext
Loop
    fila = fila + 1
    If chesub.Value = 1 Then
      FrmCaa.GridK.Rows = FrmCaa.GridK.Rows + 1
      GridK.TextMatrix(fila, 5) = "Doc.:" & LQ_TIPDOC '& " " & LQ_MONEDA
      GridK.TextMatrix(fila, 9) = Format(SUM_TIPDOC, "0.00")
    End If
    SUM_TIPDOC = 0

    fila = fila + 1
    FrmCaa.GridK.Rows = FrmCaa.GridK.Rows + 1
    If chesub.Value = 1 Then
      If LQ_MONEDA_2 = "D" Then
        GridK.TextMatrix(fila, 5) = "Total $."
      Else
        GridK.TextMatrix(fila, 5) = "Total S/."
      End If
      GridK.TextMatrix(fila, 9) = Format(SUM_MONEDA, "0.00")
    End If
    SUM_MONEDA = 0
     
     
   
     
  'fila = fila + 1
  'FrmCaa.GridK.Rows = FrmCaa.GridK.Rows + 1
  
 ' GridK.TextMatrix(fila, 2) = " "
 ' GridK.TextMatrix(fila, 3) = "S A L D O = "
 ' GridK.Row = fila
 ' GridK.Col = 3
 ' FrmCaa.GridK.CellFontBold = True
 ' GridK.TextMatrix(fila, 9) = Format(cSaldo, "0.00")
 ' GridK.Col = 4
 ' FrmCaa.GridK.CellFontBold = True
  
 
If fila = 0 Then
   pu_flag = 1
   Screen.MousePointer = 0
End If
SIGUE:

End Sub

Public Sub DETALLE()
Dim cIngreso As Currency
Dim cSalida As Currency
Dim cSaldo As Currency
Dim fila1 As Integer
Dim PRECIO As Currency
Dim cuenta As String * 6
Dim articulo As String * 6
Dim FINAL As String * 1
Dim Num_Fac As String * 5
Dim Num_Ser As String * 5
Dim WW_CODART As Long
Dim wFAR_CODART
Dim fila As Integer
Dim WS_SALDO As Currency
Dim WS_SALDO2  As Currency
Dim Band As String * 1
Dim xCODCIA As String
Dim vCONCEPTO, vdocum
FrmCaa.GRIDG.Cols = 15
FrmCaa.GRIDG.Clear
FrmCaa.GRIDG.ColWidth(0) = 1000
FrmCaa.GRIDG.ColWidth(1) = 2000
FrmCaa.GRIDG.ColWidth(2) = 1000
FrmCaa.GRIDG.ColWidth(3) = 1500
FrmCaa.GRIDG.ColWidth(4) = 1200
FrmCaa.GRIDG.ColWidth(5) = 0
FrmCaa.GRIDG.ColWidth(6) = 0
FrmCaa.GRIDG.ColWidth(7) = 0
FrmCaa.GRIDG.ColWidth(8) = 0
FrmCaa.GRIDG.ColWidth(9) = 0
FrmCaa.GRIDG.ColWidth(10) = 0
FrmCaa.GRIDG.ColWidth(11) = 1
FrmCaa.GRIDG.ColWidth(12) = 1
FrmCaa.GRIDG.ColWidth(13) = 1
FrmCaa.GRIDG.ColWidth(14) = 1

FrmCaa.GRIDG.TextMatrix(0, 0) = "Fecha"
FrmCaa.GRIDG.TextMatrix(0, 1) = "Tip.Doc."
FrmCaa.GRIDG.TextMatrix(0, 2) = "Documento"
FrmCaa.GRIDG.TextMatrix(0, 3) = "Glosa"
FrmCaa.GRIDG.TextMatrix(0, 4) = "Importe"
FrmCaa.GRIDG.TextMatrix(0, 5) = "Salida"
FrmCaa.GRIDG.TextMatrix(0, 6) = "Saldo"
FrmCaa.GRIDG.TextMatrix(0, 7) = "Fec.Vcto."
FrmCaa.GRIDG.TextMatrix(0, 8) = "Vendedor"
FrmCaa.GRIDG.TextMatrix(0, 9) = "Usuario"
FrmCaa.GRIDG.TextMatrix(0, 10) = "Hora "

xCODCIA = LK_CODCIA
Screen.MousePointer = 11
pu_fecha = #1/1/1990#
FINAL = "S"
'xCODCIA = "01"
fila = 0
'Do Until wcar_llave.EOF
'    wc_acuenta = wc_acuenta + Val(wcar_llave!mov_importe)
'    wcar_llave.MoveNext
'Loop

wcar_llave.Requery
If wcar_llave.EOF Then
   Screen.MousePointer = 0
   MsgBox "No ha efectuado Pagos.", 48, Pub_Titulo
   pu_flag = 1
   GoTo SIGUE
End If
FrmCaa.GRIDG.Rows = 1
Do Until wcar_llave.EOF
  fila = fila + 1
  FrmCaa.GRIDG.Rows = FrmCaa.GRIDG.Rows + 1
  GRIDG.TextMatrix(fila, 0) = Format(wcar_llave!MOV_FECHA, "dd/mm/yy")
  GRIDG.TextMatrix(fila, 1) = Trim(wcar_llave!MOV_SUNAT)
  vdocum = wcar_llave!MOV_serie & "-" & wcar_llave!MOV_numfac
  vCONCEPTO = Trim(wcar_llave!MOV_GLOSA)
  GRIDG.TextMatrix(fila, 2) = vdocum
  GRIDG.TextMatrix(fila, 3) = vCONCEPTO
  GRIDG.TextMatrix(fila, 4) = "0.00"
  GRIDG.TextMatrix(fila, 5) = "0.00"
  'If caa_menor!CAA_IMPORTE > 0 Then
     GRIDG.TextMatrix(fila, 4) = wcar_llave!MOV_IMPORTE
 '    WS_SALDO = WS_SALDO + Val(caa_menor!CAA_IMPORTE)
 ' Else
 '   GRIDG.TextMatrix(fila, 5) = caa_menor!CAA_IMPORTE
 '   WS_SALDO = WS_SALDO + Val(caa_menor!CAA_IMPORTE)
 ' End If
  
  'GRIDG.TextMatrix(fila, 6) = WS_SALDO
  'GRIDG.TextMatrix(fila, 7) = caa_menor!CAA_FECHA_VCTO
  'GRIDG.TextMatrix(fila, 8) = Nulo_Valor0(caa_menor!CAA_CODVEN)
  'GRIDG.TextMatrix(fila, 9) = Nulo_Valors(caa_menor!CAA_CODUSU)
  'GRIDG.TextMatrix(fila, 10) = DatePart("h", caa_menor!CAA_hora) & ":" & DatePart("n", caa_menor!CAA_hora)
  
  
  wcar_llave.MoveNext
Loop
  If WS_SALDO <> pu_saldo Then
     MsgBox "..."
  End If
  
SIGUE:
Screen.MousePointer = 1

End Sub



Public Sub ACU_GRAF()
Dim aFECHA()
Dim aSALDO()
Dim aDIAS()
Dim aPRO()
Dim xCODCIA
Dim cuenta
Dim wfecha
Dim WDIAS
Dim tra, i, wcupro
Dim wpro As Currency
Dim WPRO_TEM As Currency
Dim wranH
xCODCIA = "SELECT * FROM CARACU WHERE  CAA_CODCIA = ? AND CAA_CP = ?  AND CAA_CODCLIE = ? AND CAA_FECHA >= ? AND CAA_FECHA <= ?  ORDER BY CAA_CP, CAA_CODCLIE, CAA_CODCIA, CAA_FECHA, CAA_NUM_OPER"
Set wPSCAR_MAYOR = CN.CreateQuery("", xCODCIA)
wPSCAR_MAYOR(0) = " "
wPSCAR_MAYOR(1) = " "
wPSCAR_MAYOR(2) = 0
wPSCAR_MAYOR(3) = LK_FECHA_DIA
wPSCAR_MAYOR(4) = LK_FECHA_DIA
Set wcar_mayor = wPSCAR_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
Screen.MousePointer = 11
wPSCAR_MAYOR(0) = LK_CODCIA
wPSCAR_MAYOR(1) = pu_cp
wPSCAR_MAYOR(2) = Val(i_codcli.Text)
wPSCAR_MAYOR(3) = #1/1/1998# 'pu_fecha
wPSCAR_MAYOR(4) = #1/1/1999# 'pu_fecha
wcar_mayor.Requery
pu_flag = 0
If wcar_mayor.EOF Then
   Screen.MousePointer = 0
   pu_flag = 1
   GoTo SIGUE
End If
ReDim aFECHA(wcar_mayor.RowCount)
ReDim aSALDO(wcar_mayor.RowCount)
ReDim aDIAS(wcar_mayor.RowCount)
ReDim aPRO(wcar_mayor.RowCount)
wcupro = 0
wpro = 0
WPRO_TEM = 0
cuenta = 1
WDIAS = 0
wfecha = wcar_mayor!CAA_FECHA
If wcar_mayor!CAA_SIGNO_CAR = 1 Then
'  WPRO = WPRO + wcar_mayor!CAA_IMPORTE
End If
tra = "T"
Do Until wcar_mayor.EOF
If wfecha <> wcar_mayor!CAA_FECHA Then
   WPRO_TEM = wpro / wcupro
   aPRO(cuenta) = WPRO_TEM
   cuenta = cuenta + 1
   aFECHA(cuenta) = wcar_mayor!CAA_FECHA
   aSALDO(cuenta) = wcar_mayor!CAA_SALDO
   WDIAS = DateDiff("d", wfecha, wcar_mayor!CAA_FECHA)
   aDIAS(cuenta) = WDIAS
   wfecha = wcar_mayor!CAA_FECHA
   tra = "T"
   wpro = 0
   If wcar_mayor!CAA_SIGNO_CAR = 1 Then
       wpro = wpro + wcar_mayor!CAA_IMPORTE
   End If
   wcupro = 1
Else
   aFECHA(cuenta) = wcar_mayor!CAA_FECHA
   aSALDO(cuenta) = wcar_mayor!CAA_SALDO
   aDIAS(cuenta) = WDIAS
   If wcupro = 0 Then
     WPRO_TEM = wpro / 1
   Else
     WPRO_TEM = wpro / wcupro
   End If
   
   aPRO(cuenta) = WPRO_TEM
   tra = " "
   If wcar_mayor!CAA_SIGNO_CAR = 1 Then
       wpro = wpro + wcar_mayor!CAA_IMPORTE
       wcupro = wcupro + 1
   End If

End If
wcar_mayor.MoveNext
Loop
If tra = "T" Then
'   aFECHA(cuenta) = wcar_mayor!caa_fecha
'   aSALDO(cuenta) = wcar_mayor!caa_saldo
'   wdias = DateDiff("d", wfecha, wcar_mayor!caa_fecha)
'   aDIAS(cuenta) = wdias
'   wfecha = wcar_mayor!caa_fecha
End If

'For i = 1 To cuenta
'  msgbox i & "  :" & aFECHA(i) & " - " & aSALDO(i) & " - " & aDIAS(i)
'Next i

  
  'Dim i As Integer
  Dim xlchart As Chart
  Dim wranF, wranE, wranG
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  xl.Application.Visible = True
  xl.Workbooks.Open ("C:\Mis documentos\pruebas2.xls")
  Set xlchart = Charts(1)
  xlchart.Activate
  ActiveChart.SeriesCollection(1).Formula = "=SERIES(Hoja1!$F$9,Hoja1!$E$10:$E$" & 9 + cuenta & ",Hoja1!$F$10:$F$" & 9 + cuenta & ",1)"
  ActiveChart.SeriesCollection(2).Formula = "=SERIES(Hoja1!$G$9,Hoja1!$E$10:$E$" & 9 + cuenta & ",Hoja1!$G$10:$G$" & 9 + cuenta & ",2)"
  'ActiveChart.SeriesCollection(3).Formula = "=SERIES(Hoja1!$H$9,Hoja1!$E$" & 9 + CUENTA & ",Hoja1!$H$" & 9 + CUENTA & ",3)"
  ActiveChart.SeriesCollection(3).Formula = "=SERIES(Hoja1!$H$9,Hoja1!$E$10:$E$" & 9 + cuenta & ",Hoja1!$H$10:$H$" & 9 + cuenta & ",3)"
  'ActiveChart. = 'i_nomcli.Caption
  ActiveChart.Refresh

  Worksheets(1).Visible = True
  Worksheets(1).Activate

   For i = 1 To cuenta
     wranE = "e" & (9 + i)
     wranF = "f" & (9 + i)
     wranG = "g" & (9 + i)
     wranH = "h" & (9 + i)
     xl.Range(wranE).Value = aFECHA(i)
     xl.Range(wranF).Value = aSALDO(i)
     xl.Range(wranG).Value = aDIAS(i)
     xl.Range(wranH).Value = aPRO(i)
     
   Next i
   xlchart.Activate
 

Screen.MousePointer = 0

SIGUE:
End Sub
