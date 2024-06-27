VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmCaa 
   Caption         =   "Cuenta Corriente de Clientes o Proveedores"
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
   Begin ComctlLib.ListView LV_CLI 
      Height          =   855
      Left            =   7560
      TabIndex        =   25
      Tag             =   "0"
      Top             =   5760
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
   Begin VB.CommandButton cmdchequeo 
      Caption         =   "Chequeo de Saldos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   625
      Left            =   10440
      TabIndex        =   34
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Estado de Cuenta Grafico en Excel "
      Height          =   375
      Left            =   600
      TabIndex        =   24
      Top             =   5400
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid GridK 
      Height          =   4305
      Left            =   0
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7594
      _Version        =   393216
      Rows            =   3
      BackColor       =   16777215
      BackColorFixed  =   12632256
      BackColorBkg    =   16777215
      GridColorFixed  =   8421504
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
      Height          =   4305
      Left            =   0
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7594
      _Version        =   393216
      BackColorFixed  =   12632256
      BackColorBkg    =   16777215
      GridColorFixed  =   8421504
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
   Begin VB.Frame Consulta1 
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   -75
      Width           =   10215
      Begin VB.CommandButton cmdlimcre 
         Height          =   375
         Left            =   9360
         Picture         =   "FrmCaa.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H008B4914&
         Caption         =   "Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox i_codcli 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   0
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H008B4914&
         Caption         =   "Proveedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Index           =   1
         Left            =   3495
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   220
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Consultar :"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre/Razon Social:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   3120
         TabIndex        =   30
         Top             =   720
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.Label i_limcre 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   7320
         TabIndex        =   27
         Top             =   240
         Width           =   1905
         WordWrap        =   -1  'True
      End
      Begin VB.Label l1 
         Caption         =   "Limite Credito:"
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
         Left            =   5880
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label i_nomCLI 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   3120
         TabIndex        =   10
         Top             =   720
         Width           =   6840
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton CmdImprimir 
      BackColor       =   &H00FFFFFF&
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
      Height          =   795
      Left            =   10515
      Picture         =   "FrmCaa.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3255
      Width           =   1110
   End
   Begin VB.CommandButton SALIR 
      BackColor       =   &H00FFFFFF&
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
      Height          =   795
      Left            =   10530
      Picture         =   "FrmCaa.frx":1444
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4485
      Width           =   1110
   End
   Begin RichTextLib.RichTextBox DOCUM2 
      Height          =   255
      Left            =   7800
      TabIndex        =   6
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"FrmCaa.frx":1CBA
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
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"FrmCaa.frx":1D30
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
      Left            =   360
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   0
      TabIndex        =   11
      Top             =   960
      Width           =   10215
      Begin VB.CommandButton cmddocu 
         Caption         =   "&Mostrar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8520
         TabIndex        =   37
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdProcesa 
         Caption         =   "&Mostrar"
         Height          =   375
         Left            =   3600
         TabIndex        =   36
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton OPCONSUL 
         Caption         =   "Cuentas Detallada"
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Caption         =   "Opciones"
         Height          =   855
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   4935
         Begin VB.CheckBox Agrupar 
            Caption         =   "Mostrar Extornos."
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox ultimo 
            Caption         =   "Imp.Ultimos"
            Height          =   375
            Left            =   2160
            TabIndex        =   21
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtnum 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3360
            MaxLength       =   2
            TabIndex        =   20
            Top             =   360
            Visible         =   0   'False
            Width           =   375
         End
      End
      Begin VB.OptionButton OPCONSUL 
         Caption         =   "Pendientes."
         Height          =   375
         Index           =   1
         Left            =   5400
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Documentos"
         Height          =   855
         Left            =   5160
         TabIndex        =   14
         Top             =   600
         Width           =   4935
         Begin VB.CheckBox chesub 
            Caption         =   "Totales por Tipo de Documento."
            Height          =   255
            Left            =   600
            TabIndex        =   29
            Top             =   550
            Value           =   1  'Checked
            Width           =   3255
         End
         Begin VB.OptionButton OPCONSUL2 
            Caption         =   "Ambos."
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   17
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.OptionButton OPCONSUL2 
            Caption         =   "Cancelados"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   16
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton OPCONSUL2 
            Caption         =   "Activos"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Ver Documento :"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Lista de Documentos:"
      Height          =   255
      Left            =   0
      TabIndex        =   38
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4914&
      BorderStyle     =   1  'Fixed Single
      Height          =   7095
      Index           =   5
      Left            =   10320
      TabIndex        =   31
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label momento 
      Alignment       =   2  'Center
      Caption         =   "Procesando Información. un momento..."
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
      Left            =   2640
      TabIndex        =   22
      Top             =   3120
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
      TabIndex        =   2
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
Dim pu_zona As Long
Dim pu_subzona As Long
Dim pu_ruc As String * 12
Dim wcar_mayor As rdoResultset
Dim wPSCAR_MAYOR As rdoQuery
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
'xl.Cells(6, 1) = "FECHA APROB: " & cli_llave!CLI_FECHA_APROB
xl.Cells(3, 1) = "'" & Format(LK_FECHA_DIA, "dd/mm/yyyy")
xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))

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



  For J = 0 To 11
     wranF = LETRAS(J) & ":" & LETRAS(J)
     If GridK.ColWidth(J) < 100 Then
        xl.Columns(wranF).ColumnWidth = 0
     Else
        xl.Range(wranF).Select
        xl.Columns(wranF).ColumnWidth = GridK.ColWidth(J) / 80
     End If
  Next J


GoSub LETRAS

wranF = "A8:" & "L8"
xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 3

If Left(xl.Cells(7, 11), 4) = "Mont" Then
   wranF = "$A$1:$K$" & i + 7
   xl.Range(wranF).Select
   xl.ActiveSheet.PageSetup.PrintArea = wranF
Else
   wranF = "$A$1:$I$" & i + 7
   xl.Range(wranF).Select
   xl.ActiveSheet.PageSetup.PrintArea = wranF
'   xl.ActiveWindow.SelectedSheets.PrintOut Copies:=1
End If

xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
xl.Cells(2, 1) = "CONSULTA DE CTAS. CTES. "
xl.Cells(3, 1) = "'" & Format(LK_FECHA_DIA, "dd/mm/yyyy")
xl.DisplayAlerts = False
'xl.Worksheets(1).Protect PUB_CLAVE  quitado gts
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
  xl.Workbooks.Open Left(Trim(PUB_RUTA_OTRO), 1) & ":\ADMIN\STANDAR\KARDEX2.xls", 0, True, 4
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



Private Sub cmdchequeo_Click()
Dim CONTADOR As Integer
Dim WS_SALDO As Currency
Dim WS_SALDO2 As Currency
Dim ws_saldo_caa As Currency
Dim ART, CAA As rdoResultset
Dim PSCAA As rdoQuery
Dim fx As Integer

pub_cadena = "SELECT * FROM CARACU WHERE CAA_CP=?  AND CAA_CODCLIE=? AND CAA_CODCIA=? ORDER BY CAA_CP, CAA_CODCLIE, CAA_CODCIA, CAA_FECHA, CAA_NUM_OPER"
Set PSCAA = CN.CreateQuery("", pub_cadena)
PSCAA.rdoParameters(0) = " "
PSCAA.rdoParameters(1) = 0
PSCAA.rdoParameters(2) = " "
Set CAA = PSCAA.OpenResultset(rdOpenKeyset, rdConcurValues)

Label1.Visible = True
DoEvents
CONTADOR = 0
If option1(0).Value Then
  pu_cp = "C"
Else
  pu_cp = "P"
End If
GoSub HISTORIA_CAA
Label1.Visible = False
DoEvents
MsgBox "Proceso Terminado Satisfactoriamente...", 48, Pub_Titulo
Exit Sub
SALIDA_ERROR:

fin:
Exit Sub

HISTORIA_CAA:
On Error GoTo PASE
'Kill PUB_RUTA_OTRO & "WSLOG.txt"
On Error GoTo 0
pu_codclie = 0
pu_codcia = LK_CODCIA
SQ_OPER = 2
fx = 0
LEER_CLI_LLAVE
If cli_mayor.EOF = True Then Return
Do Until cli_mayor.EOF
   PSCAA.rdoParameters(0) = pu_cp
   PSCAA.rdoParameters(1) = cli_mayor!cli_codclie
   PSCAA.rdoParameters(2) = cli_mayor!CLI_CODCIA
   Label1.Caption = cli_mayor!cli_codclie & " " & cli_mayor.RowCount & "/ " & cli_mayor.AbsolutePosition
   DoEvents
   fx = 0
   CONTADOR = 0
REGRESA:
   CAA.Requery
   WS_SALDO = 0
   ws_saldo_caa = 0
   If fx = 1 Then
'    WRITE_LOG pu_codcia & "-" & pu_cp & " - " & cli_mayor!CLI_CODCLIE
   End If
   If CAA.EOF = False Then
      WS_SALDO = 0
      Do Until CAA.EOF
         WS_SALDO = WS_SALDO + val(CAA!CAA_IMPORTE)
         WS_SALDO = redondea(WS_SALDO)
         ws_saldo_caa = val(CAA!CAA_SALDO)
         If fx = 1 Then
            CAA.Edit
            CAA!CAA_SALDO = WS_SALDO
            CAA.Update
         End If
         CAA.MoveNext
      Loop
      CONTADOR = CONTADOR + 1
      WS_SALDO2 = CAR_TOT_CPX(pu_cp, pu_codcia, cli_mayor!cli_codclie)
      'WS_SALDO2 = CAR_TOT_CPX(pu_cp, pu_codcia, 192662) ' opcional para detalle
      If CONTADOR >= 4 Then
         MsgBox "Anotar Codigo .....Revisar este Cliente ver con detalle (ojo) : " & cli_mayor!cli_codclie
         GoTo OTRO
      End If
      If val(WS_SALDO) = val(ws_saldo_caa) And val(WS_SALDO2) = val(WS_SALDO) Then
      Else
        If fx = 1 Then
            GoTo REGRESA
        Else
            fx = 1
        End If
      GoTo REGRESA
      End If
 End If
OTRO:
'  Label1.Caption = cli_mayor!cli_nombre
  DoEvents
  cli_mayor.MoveNext
Loop

Return
Exit Sub
PASE:
Resume Next

End Sub

Private Sub cmddocu_Click()
If val(i_codcli.Text) <= 0 Then
 Exit Sub
End If
If GRIDG.Visible And GRIDG.Rows > 1 Then
 GRIDG.Visible = False
 GridK.Visible = True
 GridK.SetFocus
 Exit Sub
End If

'If Trim(i_fecha.text) <> "" Then
'    If Not IsDate(i_fecha.text) Then
'       MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
'       Exit Sub
'    End If
'    pu_fecha = i_fecha.text
'Else
'  pu_fecha = #1/1/1900#

'End If
'GridK.Visible = True
GRIDG.Visible = False
pu_flag = 0
Screen.MousePointer = 11
GridK.Visible = False
DoEvents
momento.Visible = True
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
momento.Visible = False
DoEvents

If pu_flag = 0 Then
 GridK.Visible = True
 GRIDG.Visible = False
 GridK.Row = 1
 GridK.COL = 1
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
If val(i_codcli.Text) = 0 Then i_codcli.SetFocus: Exit Sub
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
PSLIM(1) = val(i_codcli.Text)
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
momento.Visible = True
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
GRIDG.COL = 1
GRIDG.Row = 1
GRIDG.SetFocus
momento.Visible = False
DoEvents
Exit Sub
OJO:
MsgBox "Intente Nuevamente.", 48, Pub_Titulo
End Sub

Private Sub CmdProcesa_Click()
If FrmCaa.i_codcli.Text = "" Then
  MsgBox " Codigo No Existe..", 48, Pub_Titulo
  FrmCaa.i_codcli.SetFocus
  Exit Sub
End If
'If Trim(i_fecha.text) <> "" Then
'    If Not IsDate(i_fecha.text) Then
'       MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
'       Exit Sub
'    End If
'    pu_fecha = i_fecha.text
'Else
'  pu_fecha = #1/1/90#'

'End If
pu_ultimo = False
pu_maximo = 99
If ultimo.Value = 1 Then
   If val(txtnum.Text) <= 0 Then
     MsgBox "No procede ..", 48, Pub_Titulo
     txtnum.SetFocus
      Exit Sub
   End If
   pu_fecha = #1/1/1900#
   pu_ultimo = True
End If
If Agrupar.Value = 1 Then
  pu_maximo = 0
End If
pu_titulo = "  E S T A D O    D E    C U E N T A    "
Screen.MousePointer = 11
GRIDG.Visible = False
DoEvents
momento.Visible = True
DoEvents
GENERAL_CAA
'Exit Sub
momento.Visible = False
DoEvents

If pu_flag = 1 Then
 GRIDG.Visible = False
 MsgBox "NO Tiene estado de Cuenta desde esta fecha :" & pu_fecha, 48, Pub_Titulo
 FrmCaa.i_codcli.SetFocus
ElseIf pu_flag = 3 Then
 GRIDG.Visible = False
 MsgBox "Intente Nuevamente...."
Else
  GRIDG.Visible = True
  GRIDG.COL = 2
  GRIDG.Row = 1
  GRIDG.SetFocus
End If
Screen.MousePointer = 0
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

If OPCONSUL(1).Value = True Or option1(1).Value Then
  Exit Sub
End If
 Dim WSUM As Currency
 Dim TEXTO
 TEXTO = ""
 Dim PB
 PB = Chr(10) & Chr(13)
 If GRIDG.COL = 2 Then
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
       TEXTO = TEXTO + Left(art_LLAVE!art_nombre, 20) & " : " & FFF_menor!far_descri & " : " & Format((FFF_menor!far_cantidad / FFF_menor!FAR_equiv), "0.00") + PB
       WSUM = WSUM + FFF_menor!far_cantidad / FFF_menor!FAR_equiv
       FFF_menor.MoveNext
    Loop
    MsgBox "Contenido de Celda : " + PB + TEXTO + PB + "Total =   " + Format(WSUM, "0.00"), vbInformation, Pub_Titulo
    End If
    GRIDG.SetFocus
    Exit Sub
End If
 If GRIDG.COL = 8 Then
    PUB_CODVEN = GRIDG.TextMatrix(GRIDG.Row, 8)
    pu_codcia = LK_CODCIA
    SQ_OPER = 1
    LEER_VEN_LLAVE
    If Not ven_llave.EOF Then
    TEXTO = ven_llave!VEM_NOMBRE
    MsgBox "Contenido de Celda : " + PB + TEXTO, vbInformation, Pub_Titulo
    GRIDG.SetFocus
    End If
    Exit Sub
End If
If GRIDG.COL = 10 Then
 MsgBox "Contenido de Celda : " + PB + Format(GRIDG.Text, "hh:mm:ss AMPM"), 48, Pub_Titulo
 GRIDG.SetFocus
 Exit Sub
End If
If GRIDG.COL = 9 Then
 usu.Requery
 If usu.EOF Then
   MsgBox "Error de Usuarios Vuelva a Ingresar al Sistema", vbCritical, Pub_Titulo
   End
 End If
 Do Until usu.EOF
  If Trim(usu!USU_KEY) = Trim(GRIDG.Text) Then
    MsgBox "Contenido de Celda : " + PB + Trim(usu!USU_NOMBRE), vbInformation, Pub_Titulo
    GRIDG.SetFocus
    Exit Sub
  End If
  usu.MoveNext
 Loop
 GRIDG.SetFocus
 Exit Sub
End If

MsgBox "Contenido de Celda : " + PB + GRIDG.TextMatrix(GRIDG.Row, GRIDG.COL), vbInformation, Pub_Titulo
GRIDG.SetFocus

End Sub

Private Sub GridK_DblClick()
 Dim pub_mensajeText
 Dim PB
 PB = Chr(10) & Chr(13) & Chr(10) & Chr(13)
 MsgBox "Contenido de Celda : " + PB + GridK.TextMatrix(GridK.Row, GridK.COL), vbInformation, Pub_Titulo
End Sub

Private Sub i_codcli_Change()
    pu_codclie = val(i_codcli.Text)
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
pu_fecha = #1/1/1990#
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


FrmCaa.GridK.COL = 1
FrmCaa.GridK.ColAlignment(1) = 0
FrmCaa.GridK.COL = 2
FrmCaa.GridK.ColAlignment(2) = 0
FrmCaa.GridK.COL = 3
FrmCaa.GridK.ColAlignment(3) = 1
FrmCaa.GridK.COL = 4
FrmCaa.GridK.ColAlignment(4) = 2
FrmCaa.GridK.COL = 5
FrmCaa.GridK.ColAlignment(5) = 1
fecha = DateAdd("m", -1, LK_FECHA_DIA)
fech = Format(fecha, "mm")
'i_fecha.text = "01" & "/" & fech & "/" & Right(Format(fecha, "yyyy"), 4)
LV_CLI.Width = 3000
LV_CLI.Height = 3000
LV_CLI.Top = 1000
LV_CLI.Left = 3000
FrmCaa.Agrupar.Value = 0
OPCONSUL(0).Value = False
OPCONSUL(1).Value = True
OPCONSUL_Click 1
i_codcli.TabIndex = 0
ultimo.Enabled = False
End Sub


Private Sub GRIDK_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  FrmCaa.GridK.Visible = False
  FrmCaa.i_codcli.SetFocus
End If

If KeyAscii = 13 Then
 'GridK.Visible = False
  If Trim(GridK.TextMatrix(GridK.Row, 8)) = "" Then
  Exit Sub
  End If
 pu_numdoc = val(GridK.TextMatrix(GridK.Row, 8))
 PUB_SERDOC = val(GridK.TextMatrix(GridK.Row, 7))
 pu_saldo = val(GridK.TextMatrix(GridK.Row, 9))
 PUB_TIPDOC = GridK.TextMatrix(GridK.Row, 2)
 
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
    If option1(0).Value Then
      option1(0).Value = False
      option1(1).Value = True
      Option1_Click (1)
    Else
     option1(0).Value = True
      option1(1).Value = False
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
  If loc_key > LV_CLI.ListItems.count Then loc_key = LV_CLI.ListItems.count
  GoTo posicion
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo posicion
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > LV_CLI.ListItems.count Then loc_key = LV_CLI.ListItems.count
 GoTo posicion
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo posicion
End If
GoTo fin
posicion:
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
Dim VAR
If Len(i_codcli.Text) = 0 Or IsNumeric(i_codcli.Text) = True Then
   LV_CLI.Visible = False
   Exit Sub
End If
If LV_CLI.Visible = False And KeyCode <> 13 Or Len(i_codcli.Text) = 1 Then
    VAR = Asc(i_codcli.Text)
    VAR = VAR + 1
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    numarchi = 1
   ' archi = "SELECT CLI_CODCLIE , CLI_CODCIA, CLI_CP, CLI_NOMBRE , CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM  FROM CLIENTES WHERE CLI_CP = '" & pu_cp & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & i_codcli.Text & "' AND  '" & VAR & "' ORDER BY CLI_NOMBRE"
    archi = "SELECT CLI_CODCLIE , CLI_CODCIA, CLI_CP, CLI_NOMBRE, CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM, TAB_NOMLARGO  FROM CLIENTES,TABLAS WHERE (TAB_CODCIA = '00') AND (TAB_TIPREG = 35) AND (TAB_NUMTAB = CLI_ZONA_NEW) AND CLI_CP = '" & pu_cp & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & i_codcli.Text & "' AND  '" & VAR & "' ORDER BY CLI_NOMBRE"
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
   If loc_key + 8 > LV_CLI.ListItems.count Then
      LV_CLI.ListItems.Item(LV_CLI.ListItems.count).EnsureVisible
   Else
     LV_CLI.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If


End Sub

Private Sub i_codcli_KeyPress(KeyAscii As Integer)
Dim VAR As String
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As ListItem    ' Variable FoundItem.
FrmCaa.i_nomCLI.Caption = ""
If FrmCaa.option1(0).Value Then
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
If KeyAscii = 13 And Left(i_codcli.Text, 1) = "+" Then GoTo buscar

pu_codclie = val(i_codcli.Text)
If Len(i_codcli.Text) = 0 Then
   Exit Sub
End If
If pu_codclie <> 0 And IsNumeric(i_codcli.Text) = True Then
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
        i_codcli.Text = cli_ruc!cli_codclie
   End If
   SQ_OPER = 1
   pu_codclie = val(i_codcli.Text)
   pu_codcia = LK_CODCIA
   LEER_CLI_LLAVE
   If cli_llave.EOF Then
     Azul i_codcli, i_codcli
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     GoTo fin
   Else
      i_nomCLI.Caption = Trim(cli_llave(3)) & " - " & "RUC. " & cli_llave!cli_ruc_esposo
   End If
Else
   If loc_key > LV_CLI.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   valor = UCase(LV_CLI.ListItems.Item(loc_key).Text)
   If Trim(UCase(i_codcli.Text)) = Left(valor, Len(Trim(i_codcli.Text))) Then
   Else
      Exit Sub
   End If
   
   
   i_codcli.Text = Trim(LV_CLI.ListItems.Item(loc_key).SubItems(1))
   pu_codclie = val(i_codcli.Text)
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_CLI_LLAVE
End If
LV_CLI.Visible = False
       FrmCaa.i_nomCLI.Caption = Trim(cli_llave(3)) & " - " & "RUC. " & cli_llave!cli_ruc_esposo
       If cli_llave!CLI_MONEDA = "S" Then
          FrmCaa.i_limcre.Caption = "  S/." & Nulo_Valor0(cli_llave!cli_limcre)
       Else
          FrmCaa.i_limcre.Caption = "  $  " & Nulo_Valor0(cli_llave!cli_limcre2)
       End If
       pu_direc = Trim(cli_llave(10))
       pu_numero = Trim(cli_llave(11))
       'pu_zona = cli_llave(12)
      ' pu_subzona = cli_llave(13)
      ' pu_ruc = cli_llave!cli_ruc_esposo
       
       If FrmCaa.CmdProcesa.Enabled Then
          FrmCaa.CmdProcesa.SetFocus
          CmdProcesa_Click
       Else
          FrmCaa.cmddocu.SetFocus
          cmddocu_Click
       End If

fin:
Exit Sub
buscar:
VAR = Mid(i_codcli.Text, 2, Len(i_codcli.Text))
numarchi = alta_vista_nombre(LV_CLI, VAR, pu_cp)
If numarchi = 0 Then
  LV_CLI.Visible = False
  MsgBox "Alta Vista: No Existe .. Esta descripcion..", 48, Pub_Titulo
Else
  LV_CLI.Visible = True
  i_codcli.SetFocus
End If
loc_key = 1

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
    OPCONSUL(1).Enabled = True
    'ultimo.Enabled = True
    'Agrupar.Value = 1
    txtnum.Enabled = True
    OPCONSUL2(0).Enabled = False
    OPCONSUL2(1).Enabled = False
    cmddocu.Enabled = False
    CmdProcesa.Enabled = True
    CmdProcesa_Click
ElseIf Index = 1 Then
    CmdProcesa.Enabled = False
    ultimo.Enabled = False
    txtnum.Enabled = False
    OPCONSUL(0).Enabled = True
    OPCONSUL2(0).Enabled = True
    OPCONSUL2(1).Enabled = True
    OPCONSUL2(0).Value = True
    cmddocu.Enabled = True
    cmddocu_Click
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

Public Sub DOCUMENTO(WDOCU As Integer)
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
If LK_FLAG_SOS = "A" Then
  If OPCONSUL2(1).Value Then
     xCODCIA = "SELECT * FROM CARTERA WHERE  CAR_CODCIA = ? AND CAR_CP = ? AND CAR_CODCLIE = ?   AND CAR_FLAG_SO = 'A' AND CAR_IMPORTE = 0 ORDER BY CAR_CODCIA, CAR_CP, CAR_CODCLIE, CAR_MONEDA , CAR_TIPDOC, CAR_NUMGUIA , CAR_FECHA_INGR"
  ElseIf OPCONSUL2(0).Value Then
     xCODCIA = "SELECT * FROM CARTERA WHERE  CAR_CODCIA = ? AND CAR_CP = ? AND CAR_CODCLIE = ?   AND CAR_FLAG_SO = 'A' AND CAR_IMPORTE <> 0 ORDER BY CAR_CODCIA, CAR_CP, CAR_CODCLIE, CAR_MONEDA , CAR_TIPDOC, CAR_NUMGUIA , CAR_FECHA_INGR"
  ElseIf OPCONSUL2(2).Value Then
     xCODCIA = "SELECT * FROM CARTERA WHERE  CAR_CODCIA = ? AND CAR_CP = ? AND CAR_CODCLIE = ?   AND CAR_FLAG_SO = 'A'  ORDER BY CAR_CODCIA, CAR_CP, CAR_CODCLIE, CAR_MONEDA , CAR_TIPDOC, CAR_NUMGUIA , CAR_FECHA_INGR"
  End If
Else
  If OPCONSUL2(1).Value Then
     xCODCIA = "SELECT * FROM CARTERA WHERE  CAR_CODCIA = ? AND CAR_CP = ? AND CAR_CODCLIE = ? AND CAR_IMPORTE = 0 ORDER BY CAR_CODCIA, CAR_CP, CAR_CODCLIE, CAR_MONEDA , CAR_TIPDOC, CAR_NUMGUIA , CAR_FECHA_SUNAT"
  ElseIf OPCONSUL2(0).Value Then
     xCODCIA = "SELECT * FROM CARTERA WHERE  CAR_CODCIA = ? AND CAR_CP = ? AND CAR_CODCLIE = ? AND CAR_IMPORTE <> 0 ORDER BY CAR_CODCIA, CAR_CP, CAR_CODCLIE, CAR_MONEDA , CAR_TIPDOC, CAR_NUMGUIA , CAR_FECHA_SUNAT"
  ElseIf OPCONSUL2(2).Value Then
     xCODCIA = "SELECT * FROM CARTERA WHERE  CAR_CODCIA = ? AND CAR_CP = ? AND CAR_CODCLIE = ?  ORDER BY CAR_CODCIA, CAR_CP, CAR_CODCLIE, CAR_MONEDA , CAR_TIPDOC, CAR_NUMGUIA , CAR_FECHA_SUNAT"
  End If
End If
Set wPSCAR_MAYOR = CN.CreateQuery("", xCODCIA)
wPSCAR_MAYOR.rdoParameters(0) = " "
wPSCAR_MAYOR.rdoParameters(1) = " "
wPSCAR_MAYOR.rdoParameters(2) = 0
Set wcar_mayor = wPSCAR_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
Screen.MousePointer = 11
SQ_OPER = 2
wPSCAR_MAYOR(0) = LK_CODCIA
wPSCAR_MAYOR(1) = pu_cp
wPSCAR_MAYOR(2) = val(i_codcli.Text)
wcar_mayor.Requery

If wcar_mayor.EOF Then
   pu_flag = 1
   GoTo sigue
End If

FrmCaa.GridK.Font.Size = 8
FrmCaa.GridK.Clear
FrmCaa.GridK.Cols = 15
FrmCaa.GridK.ColAlignment(4) = 1
FrmCaa.GridK.ColWidth(0) = 950
FrmCaa.GridK.ColWidth(1) = 950
FrmCaa.GridK.ColWidth(2) = 450
FrmCaa.GridK.ColWidth(3) = 1400
FrmCaa.GridK.ColWidth(4) = 2100

FrmCaa.GridK.ColWidth(5) = 950
FrmCaa.GridK.ColWidth(6) = 500
FrmCaa.GridK.ColWidth(7) = 1
FrmCaa.GridK.ColWidth(8) = 1
FrmCaa.GridK.ColWidth(9) = 1400

FrmCaa.GridK.ColWidth(10) = 1300

FrmCaa.GridK.ColWidth(11) = 1400
FrmCaa.GridK.ColWidth(12) = 2500
FrmCaa.GridK.ColWidth(13) = 900
If pu_cp = "P" Then
    FrmCaa.GridK.ColWidth(14) = 900
Else
    FrmCaa.GridK.ColWidth(14) = 0
End If

FrmCaa.GridK.TextMatrix(0, 0) = "Fec.Proc."
FrmCaa.GridK.TextMatrix(0, 1) = "Fec.Emis."
FrmCaa.GridK.TextMatrix(0, 2) = "Doc."
FrmCaa.GridK.TextMatrix(0, 3) = "Documento"
FrmCaa.GridK.TextMatrix(0, 4) = "Concepto"

FrmCaa.GridK.TextMatrix(0, 5) = "Fec.Vcto"
FrmCaa.GridK.Row = 0
FrmCaa.GridK.COL = 9
FrmCaa.GridK.CellFontBold = True
FrmCaa.GridK.CellForeColor = QBColor(12)
FrmCaa.GridK.TextMatrix(0, 9) = "Saldo Actual."
FrmCaa.GridK.TextMatrix(0, 10) = "Monto Orig."
FrmCaa.GridK.TextMatrix(0, 11) = "F.V.Orig."
FrmCaa.GridK.TextMatrix(0, 12) = " B a n c o. "
FrmCaa.GridK.TextMatrix(0, 13) = "Vend."
FrmCaa.GridK.TextMatrix(0, 14) = "Kardex"

fila = 0
FrmCaa.GridK.Rows = 1
cSaldo = 0
SUM_TIPDOC = 0
SUM_MONEDA = 0
LQ_TIPDOC = wcar_mayor!car_TIPDOC
LQ_MONEDA_2 = wcar_mayor!CAR_MONEDA
Do Until wcar_mayor.EOF
  If WDOCU = 0 Then
     'If wcar_mayor!CAR_IMPORTE <> 0 Then
     'Else
     '  GoTo valeloop
     'End If
  ElseIf WDOCU = 1 Then
     If wcar_mayor!car_importe = 0 Then
     Else
       GoTo valeloop
     End If
  End If
  fila = fila + 1
  FrmCaa.GridK.Rows = FrmCaa.GridK.Rows + 1
  If wcar_mayor!CAR_MONEDA = "D" Then
    LQ_MONEDA = "$ "
  Else
    LQ_MONEDA = "S/."
  End If
  If (LQ_TIPDOC <> wcar_mayor!car_TIPDOC Or LQ_MONEDA_2 <> wcar_mayor!CAR_MONEDA) And chesub.Value = 1 Then
     GridK.TextMatrix(fila, 5) = "Doc.:" & LQ_TIPDOC '& " " & LQ_MONEDA
     GridK.TextMatrix(fila, 9) = Format(SUM_TIPDOC, "0.000")
     fila = fila + 1
     FrmCaa.GridK.Rows = FrmCaa.GridK.Rows + 1
     SUM_TIPDOC = 0
     LQ_TIPDOC = wcar_mayor!car_TIPDOC
  End If
  If LQ_MONEDA_2 <> wcar_mayor!CAR_MONEDA And chesub.Value = 1 Then
     If LQ_MONEDA_2 = "D" Then
      GridK.TextMatrix(fila, 5) = "Total $."
     Else
      GridK.TextMatrix(fila, 5) = "Total S/."
     End If
     GridK.TextMatrix(fila, 9) = Format(SUM_MONEDA, "0.000")
     fila = fila + 1
     FrmCaa.GridK.Rows = FrmCaa.GridK.Rows + 1
     SUM_TIPDOC = 0
     SUM_MONEDA = 0
     LQ_MONEDA_2 = wcar_mayor!CAR_MONEDA
  End If
  GridK.Row = fila
  GridK.TextMatrix(fila, 0) = Format(wcar_mayor!CAR_FECHA_INGR, "dd/mm/yy")
  GridK.TextMatrix(fila, 1) = Format(wcar_mayor!CAR_FECHA_SUNAT, "dd/mm/yy")
  GridK.TextMatrix(fila, 2) = wcar_mayor!car_TIPDOC
  vCONCEPTO = Trim(wcar_mayor!car_concepto)
  If Left(Trim(wcar_mayor!car_concepto), 1) >= "0" And Left(Trim(wcar_mayor!car_concepto), 1) <= "9" Then
   vCONCEPTO = " " + Trim(wcar_mayor!car_concepto)
  End If
    If wcar_mayor!CAR_cp = "C" Then
     If Trim(wcar_mayor!car_FBG) = "F" Then
        vdocum = "Fact. " & Trim(wcar_mayor!car_NUMSER) & " - " & wcar_mayor!car_NUMFAC
     ElseIf Trim(wcar_mayor!car_FBG) = "B" Then
        vdocum = "Bolet. " & Trim(wcar_mayor!car_NUMSER) & " - " & wcar_mayor!car_NUMFAC
     ElseIf Trim(wcar_mayor!car_FBG) = "G" Then
        vdocum = "Guia . " & Trim(wcar_mayor!car_NUMSER) & " - " & Trim(wcar_mayor!car_NUMFAC)
     ElseIf Trim(wcar_mayor!car_FBG) = "N" Then
        vdocum = "N.Cred. " & wcar_mayor!car_NUMSER & "-" & wcar_mayor!car_NUMFAC
     ElseIf Trim(wcar_mayor!car_FBG) = "D" Then
        vdocum = "N.Deb. " & wcar_mayor!car_NUMSER & "-" & wcar_mayor!car_NUMFAC
     ElseIf Trim(wcar_mayor!car_FBG) = "" Then
        vdocum = "Nº:" & wcar_mayor!car_NUMSER & "-" & wcar_mayor!car_NUMFAC

     Else
        vdocum = " "
     End If
     End If
     
     If wcar_mayor!CAR_TIPMOV = 97 Then
        vdocum = "N.-Cred. " & wcar_mayor!car_NUMSER & "-" & wcar_mayor!car_NUMFAC
     End If
     If wcar_mayor!CAR_TIPMOV = 98 Then
        vdocum = "N.-Deb. " & wcar_mayor!car_NUMSER & "-" & wcar_mayor!car_NUMFAC
     End If
    
    If wcar_mayor!CAR_cp = "P" Then
       If wcar_mayor!CAR_TIPMOV = 98 Then
          vdocum = "N.-Deb. " & wcar_mayor!CAR_NUMSER_C & "-" & wcar_mayor!CAR_NUMFAC_C
       End If
       If wcar_mayor!CAR_TIPMOV = 97 Then
          vdocum = "N.-Cred. " & wcar_mayor!CAR_NUMSER_C & "-" & wcar_mayor!CAR_NUMFAC_C
       End If
       If wcar_mayor!CAR_TIPMOV = 20 Then
          If val(Nulo_Valor0(wcar_mayor!CAR_NUMFAC_C)) > 0 Then
             vdocum = "Fact. " & wcar_mayor!CAR_NUMSER_C & "-" & wcar_mayor!CAR_NUMFAC_C
          Else
             vdocum = "Guia . " & wcar_mayor!car_numguia
          End If
        End If
        If wcar_mayor!CAR_TIPMOV = 99 Then
           If val(Nulo_Valor0(wcar_mayor!CAR_NUMFAC_C)) > 0 Then
                vdocum = "Fact. " & wcar_mayor!CAR_NUMSER_C & "-" & wcar_mayor!CAR_NUMFAC_C
            Else
               vdocum = "Guia . " & Trim(wcar_mayor!car_numguia)
            End If
        End If
     Else
        If wcar_mayor!CAR_TIPMOV = 10 And LK_EMP <> "HER" And LK_EMP <> "3AA" Then
          If val(Nulo_Valor0(wcar_mayor!CAR_NUMFAC_C)) > 0 Then
             vdocum = "Fact. " & wcar_mayor!CAR_NUMSER_C & "-" & wcar_mayor!CAR_NUMFAC_C
          ElseIf Trim(wcar_mayor!car_FBG) <> "F" And Trim(wcar_mayor!car_FBG) <> "B" Then
             vdocum = "Guia . " & wcar_mayor!car_NUMSER & " - " & wcar_mayor!car_NUMFAC
          End If
        End If
     End If
     If wcar_mayor!car_TIPDOC = "LE" Then
        vdocum = " " & wcar_mayor!car_NUMSER & " - " & wcar_mayor!car_NUMFAC
     End If
    
    If wcar_mayor!CAR_TIPMOV = 0 Or wcar_mayor!car_TIPDOC = "CD" Or wcar_mayor!car_TIPDOC = "CH" Then
     ' If Val(wcar_mayor!CAR_NUM_CHEQUE) <> 0 Then
       vdocum = "Nº. " & wcar_mayor!car_NUM_CHEQUE
     ' End If
     If wcar_mayor!car_TIPDOC = "FA" And wcar_mayor!car_FBG <> " " Then
        vdocum = wcar_mayor!car_FBG & " / " & wcar_mayor!CAR_NUMSER_C & "-" & wcar_mayor!CAR_NUMFAC_C
     End If
'    If wcar_mayor!CAR_TIPDOC = "LE" Then
'       vdocum = "LETRA "
'    End If
    End If
    
  GridK.TextMatrix(fila, 3) = vdocum
  GridK.TextMatrix(fila, 4) = vCONCEPTO
  If Trim(wcar_mayor!CAR_VOUCHER) <> "" Then
    GridK.TextMatrix(fila, 4) = Nulo_Valors(wcar_mayor!car_concepto)
  End If
  GridK.TextMatrix(fila, 10) = Format(wcar_mayor!CAR_IMP_INI, "0.000")
  GridK.COL = 5
  GridK.CellAlignment = 7
  GridK.TextMatrix(fila, 5) = Format(wcar_mayor!car_fecha_vcto, "dd/mm/yy")
  
  If wcar_mayor!CAR_MONEDA = "D" Then
    GridK.TextMatrix(fila, 6) = "US$."
  Else
    GridK.TextMatrix(fila, 6) = " S/."
  End If
  GridK.TextMatrix(fila, 7) = wcar_mayor!car_SERDOC
  GridK.TextMatrix(fila, 8) = wcar_mayor!car_NUMDOC
  
  GridK.TextMatrix(fila, 9) = Format(wcar_mayor!car_importe, "0.000")
  
  GridK.TextMatrix(fila, 11) = Format(wcar_mayor!car_fecha_vcto_orig, "dd/mm/yyyy")
  GridK.TextMatrix(fila, 13) = Format(wcar_mayor!CAR_codven, "00")
  GridK.TextMatrix(fila, 14) = wcar_mayor!car_NUMSER & " - " & wcar_mayor!car_NUMFAC
  SUM_TIPDOC = SUM_TIPDOC + wcar_mayor!car_importe
  SUM_MONEDA = SUM_MONEDA + wcar_mayor!car_importe
  If Nulo_Valor0(wcar_mayor!CAR_codban) <> 0 Then
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   PUB_CODBAN = wcar_mayor!CAR_codban
   LEER_CCM_LLAVE
   If Not ccm_llave.EOF Then GridK.TextMatrix(fila, 12) = Trim(ccm_llave!CCM_NOMBRE)
  End If
  
  cSaldo = wcar_mayor!car_importe + cSaldo
valeloop:
  wcar_mayor.MoveNext
Loop
    fila = fila + 1
    If chesub.Value = 1 Then
      FrmCaa.GridK.Rows = FrmCaa.GridK.Rows + 1
      GridK.TextMatrix(fila, 5) = "Doc.:" & LQ_TIPDOC '& " " & LQ_MONEDA
      GridK.TextMatrix(fila, 9) = Format(SUM_TIPDOC, "0.000")
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
      GridK.TextMatrix(fila, 9) = Format(SUM_MONEDA, "0.000")
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
sigue:

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
FrmCaa.GRIDG.ColWidth(1) = 500
FrmCaa.GRIDG.ColWidth(2) = 1400
FrmCaa.GRIDG.ColWidth(3) = 1000
FrmCaa.GRIDG.ColWidth(4) = 1200
FrmCaa.GRIDG.ColWidth(5) = 1200
FrmCaa.GRIDG.ColWidth(6) = 1400
FrmCaa.GRIDG.ColWidth(7) = 1000
FrmCaa.GRIDG.ColWidth(8) = 600
FrmCaa.GRIDG.ColWidth(9) = 1200
FrmCaa.GRIDG.ColWidth(10) = 1000
FrmCaa.GRIDG.ColWidth(11) = 1
FrmCaa.GRIDG.ColWidth(12) = 1
FrmCaa.GRIDG.ColWidth(13) = 1
FrmCaa.GRIDG.ColWidth(14) = 1

FrmCaa.GRIDG.TextMatrix(0, 0) = "Fecha"
FrmCaa.GRIDG.TextMatrix(0, 1) = "Tipo"
FrmCaa.GRIDG.TextMatrix(0, 2) = "Documento"
FrmCaa.GRIDG.TextMatrix(0, 3) = "Concepto"
FrmCaa.GRIDG.TextMatrix(0, 4) = "Ingreso"
FrmCaa.GRIDG.TextMatrix(0, 5) = "Salida"
FrmCaa.GRIDG.TextMatrix(0, 6) = "Saldo"
FrmCaa.GRIDG.TextMatrix(0, 7) = "Fec.Vcto."
FrmCaa.GRIDG.TextMatrix(0, 8) = "Vendedor"
FrmCaa.GRIDG.TextMatrix(0, 9) = "Usuario"
FrmCaa.GRIDG.TextMatrix(0, 10) = "Hora "






xCODCIA = LK_CODCIA
Screen.MousePointer = 11

FINAL = "S"
'xCODCIA = "01"
fila = 0
    pu_codclie = val(i_codcli.Text)
    PSCAA_menor.rdoParameters(0) = pu_cp
    PSCAA_menor.rdoParameters(1) = pu_codclie
    PSCAA_menor.rdoParameters(2) = xCODCIA
    PSCAA_menor.rdoParameters(3) = pu_fecha
    PSCAA_menor.rdoParameters(5) = PUB_SERDOC
    PSCAA_menor.rdoParameters(4) = pu_numdoc
    PSCAA_menor.rdoParameters(6) = PUB_TIPDOC
    caa_menor.Requery
    If caa_menor.EOF Then
       Screen.MousePointer = 0
       MsgBox "NO existe Movimientos", 48, Pub_Titulo
       pu_flag = 1
       GoTo sigue
    End If
WS_SALDO = caa_menor!caa_SALDO_car - val(caa_menor!CAA_IMPORTE)

FrmCaa.GRIDG.Rows = 1
Do Until caa_menor.EOF
  If caa_menor!CAA_ESTADO <> "E" Then
  Else
'    GoTo sal
  End If
  fila = fila + 1
  FrmCaa.GRIDG.Rows = FrmCaa.GRIDG.Rows + 1
  'Print caa_menor!CAA_CODCLIE
  GRIDG.TextMatrix(fila, 0) = Format(caa_menor!CAA_FECHA, "dd/mm/yy")
  GRIDG.TextMatrix(fila, 1) = caa_menor!CAA_TIPDOC
    If caa_menor!CAA_CP = "C" Then
     If Trim(caa_menor!CAA_FBG) = "F" Then
        vdocum = "Fact. " & caa_menor!CAa_numser & "-" & caa_menor!CAa_numfac
     ElseIf Trim(caa_menor!CAA_FBG) = "B" Then
        vdocum = "Bolet. " & caa_menor!CAa_numser & "-" & caa_menor!CAa_numfac
     ElseIf Trim(caa_menor!CAA_FBG) = "G" Or Trim(caa_menor!CAA_FBG) = "P" Or Trim(caa_menor!CAA_FBG) = "C" Then
        vdocum = "Guia . " & caa_menor!CAa_numser & "-" & caa_menor!CAa_numfac
     ElseIf Trim(caa_menor!CAA_NOTA) = "N" And caa_menor!CAA_IMPORTE < 0 Then
        vdocum = "N.Cred." & caa_menor!caa_numser_c & "-" & caa_menor!caa_numfac_c & " del Docum:" & caa_menor!CAA_FBG & "/." & caa_menor!CAa_numser & "-" & caa_menor!CAa_numfac
     ElseIf Trim(caa_menor!CAA_NOTA) = "N" And caa_menor!CAA_IMPORTE > 0 Then
        vdocum = "N.Deb ." & caa_menor!caa_numser_c & "-" & caa_menor!caa_numfac_c
     Else
        vdocum = " "
     End If
    End If
    
     
     If caa_menor!CAA_TIPMOV = 97 Then
        vdocum = "N.-Cred. " & caa_menor!caa_numser_c & "-" & caa_menor!caa_numfac_c & " Doc.: " & caa_menor!CAA_FBG & " /" & caa_menor!CAa_numser & "-" & caa_menor!CAa_numfac
     End If
     If caa_menor!CAA_TIPMOV = 98 Then
        vdocum = "N.-Deb. " & caa_menor!caa_numser_c & "-" & caa_menor!caa_numfac_c & " Doc.: " & caa_menor!CAA_FBG & " /" & caa_menor!CAa_numser & "-" & caa_menor!CAa_numfac
     End If
    
    If caa_menor!CAA_CP = "P" Then
        If caa_menor!CAA_TIPMOV = 97 Or caa_menor!CAA_TIPMOV = 98 Then
        Else
        If val(Nulo_Valor0(caa_menor!caa_numfac_c)) > 0 Then
           vdocum = "Fact. " & caa_menor!caa_numser_c & "-" & caa_menor!caa_numfac_c
        Else
           vdocum = "Guia . " & caa_menor!CAa_numGUIA
        End If
        End If
   End If

  
  
  
  
  
  vCONCEPTO = Trim(caa_menor!caa_concepto)
  If Left(Trim(caa_menor!caa_concepto), 1) >= "0" And Left(Trim(caa_menor!caa_concepto), 1) <= "9" Then
   vCONCEPTO = " " + Trim(caa_menor!caa_concepto)
  End If
  
  GRIDG.TextMatrix(fila, 2) = vdocum
  GRIDG.TextMatrix(fila, 3) = vCONCEPTO
  GRIDG.TextMatrix(fila, 4) = "0.00"
  GRIDG.TextMatrix(fila, 5) = "0.00"
  If caa_menor!CAA_IMPORTE > 0 Then
     GRIDG.TextMatrix(fila, 4) = caa_menor!CAA_IMPORTE
     WS_SALDO = WS_SALDO + val(caa_menor!CAA_IMPORTE)
  Else
    GRIDG.TextMatrix(fila, 5) = caa_menor!CAA_IMPORTE
    WS_SALDO = WS_SALDO + val(caa_menor!CAA_IMPORTE)
  End If
  
  GRIDG.TextMatrix(fila, 6) = WS_SALDO
  GRIDG.TextMatrix(fila, 7) = caa_menor!CAA_FECHA_VCTO
  GRIDG.TextMatrix(fila, 8) = Nulo_Valor0(caa_menor!CAA_CODVEN)
  GRIDG.TextMatrix(fila, 9) = Nulo_Valors(caa_menor!CAA_CODUSU)
  GRIDG.TextMatrix(fila, 10) = DatePart("h", caa_menor!CAA_hora) & ":" & DatePart("n", caa_menor!CAA_hora)
  
sal:
  caa_menor.MoveNext
Loop
  If WS_SALDO <> pu_saldo Then
     MsgBox "..."
  End If
  
sigue:
Screen.MousePointer = 1

End Sub

Public Sub GENERAL_CAA()
Dim cIngreso As Currency
Dim cSalida As Currency
Dim cSaldo As Currency
Dim cFecha As String
Dim cTipdoc As String
Dim cFechaV As String
Dim cFecha_emi As String
Dim cvendedor As Integer
Dim ws_total As Currency
Dim chora
Dim fila1 As Integer
Dim vCONCEPTO, vdocum As String

Dim cuenta As String * 6
Dim articulo As String * 6
Dim FINAL As String * 1
Dim Num_Fac As String * 5
Dim Num_Ser As String * 5
Dim WW_CODART As Long
Dim wFAR_CODART
Dim fila, i As Integer
Dim WS_SALDO, WS_SALDO_ULT As Currency
Dim WS_SALDO2  As Currency
Dim ws_saldo_doc As Currency
Dim Band As String * 1
Dim xCODCIA As String
Dim wultimos As Integer
Dim ultimos5 As Integer
Dim fx As Integer


If LK_FLAG_SOS = "A" Then
 xCODCIA = "SELECT * FROM CARACU WHERE  CAA_CODCIA = ? AND CAA_CP = ?  AND CAA_CODCLIE = ? AND CAA_FECHA >= ? AND CAA_FLAG_SO = 'A' ORDER BY CAA_FECHA, CAA_NUM_OPER"
Else
 xCODCIA = "SELECT * FROM CARACU WHERE  CAA_CODCIA = ? AND CAA_CP = ?  AND CAA_CODCLIE = ? AND CAA_FECHA >= ?  ORDER BY CAA_FECHA, CAA_NUM_OPER"
End If
Set wPSCAR_MAYOR = CN.CreateQuery("", xCODCIA)
wPSCAR_MAYOR(0) = " "
wPSCAR_MAYOR(1) = " "
wPSCAR_MAYOR(2) = 0
wPSCAR_MAYOR(3) = LK_FECHA_DIA
Set wcar_mayor = wPSCAR_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
Screen.MousePointer = 11
wPSCAR_MAYOR(0) = LK_CODCIA
wPSCAR_MAYOR(1) = pu_cp
wPSCAR_MAYOR(2) = val(i_codcli.Text)
wPSCAR_MAYOR(3) = pu_fecha
wcar_mayor.Requery
pu_flag = 0
If wcar_mayor.EOF Then
   Screen.MousePointer = 0
   pu_flag = 1
   GoTo sigue
End If

WS_SALDO = wcar_mayor!CAA_SALDO - val(wcar_mayor!CAA_IMPORTE)


FrmCaa.GRIDG.Clear
FrmCaa.GRIDG.Font.Size = 8


FrmCaa.GRIDG.Cols = 16
FrmCaa.GRIDG.ColWidth(0) = 1000
FrmCaa.GRIDG.ColWidth(1) = 1
FrmCaa.GRIDG.ColWidth(2) = 500
FrmCaa.GRIDG.ColWidth(3) = 1500
FrmCaa.GRIDG.ColWidth(4) = 1500
FrmCaa.GRIDG.ColWidth(5) = 1200
FrmCaa.GRIDG.ColWidth(6) = 1200
FrmCaa.GRIDG.ColWidth(7) = 1400
FrmCaa.GRIDG.ColWidth(8) = 1000
FrmCaa.GRIDG.ColWidth(9) = 500
FrmCaa.GRIDG.ColWidth(10) = 1200
FrmCaa.GRIDG.ColWidth(11) = 1000
FrmCaa.GRIDG.ColWidth(12) = 1
FrmCaa.GRIDG.ColWidth(13) = 1
FrmCaa.GRIDG.ColWidth(14) = 1
FrmCaa.GRIDG.ColWidth(15) = 1

FrmCaa.GRIDG.TextMatrix(0, 0) = "Fec.Proc."
FrmCaa.GRIDG.TextMatrix(0, 1) = "Fec.Emis."
FrmCaa.GRIDG.TextMatrix(0, 2) = "Tipo"
FrmCaa.GRIDG.TextMatrix(0, 3) = "Documento"
FrmCaa.GRIDG.TextMatrix(0, 4) = "Concepto"
FrmCaa.GRIDG.TextMatrix(0, 5) = "Ingreso"
FrmCaa.GRIDG.TextMatrix(0, 6) = "Salida"
FrmCaa.GRIDG.TextMatrix(0, 7) = "Saldo"
If LK_EMP = "PAR" Then
FrmCaa.GRIDG.TextMatrix(0, 8) = "F.Cobro"
Else
FrmCaa.GRIDG.TextMatrix(0, 8) = "Fec.Vcto."
End If
FrmCaa.GRIDG.TextMatrix(0, 9) = "Vendedor"
FrmCaa.GRIDG.TextMatrix(0, 10) = "Usuario"
FrmCaa.GRIDG.TextMatrix(0, 11) = "Hora "

WS_SALDO2 = CAR_TOT_CPX(pu_cp, LK_CODCIA, val(i_codcli.Text))

fila = 0
FrmCaa.GRIDG.Rows = 1
fx = 0
ws_total = 0
ws_saldo_doc = 0


Do Until wcar_mayor.EOF
   If Agrupar.Value = 0 And wcar_mayor!CAA_ESTADO = "E" Then
   Else
   GoSub ACU
   GoSub Impri
   End If
   
   wcar_mayor.MoveNext
Loop
If fila = 0 Then
   pu_flag = 1
   Screen.MousePointer = 0
   Exit Sub
End If

fx = fila
cSaldo = WS_SALDO2

Do Until fx = 0
   GRIDG.TextMatrix(fx, 7) = Format(cSaldo, "0.00")
   If val(GRIDG.TextMatrix(fx, 5)) <> 0 Then
      cSaldo = cSaldo - GRIDG.TextMatrix(fx, 5)
   Else
      cSaldo = cSaldo - GRIDG.TextMatrix(fx, 6)
   End If
   fx = fx - 1
   
Loop
If cSaldo <> 0 Then
   Screen.MousePointer = 0
   MsgBox "..", 48, Pub_Titulo
End If


'    If fx = 1 Then
'       pu_flag = 3
'       MsgBox "Revisar...docs..."
'    End If
'cli_llave.Requery
'If redondea(WS_SALDO) <> WS_SALDO2 Or redondea(WS_SALDO) <> WS_SALDO_ULT Or cSaldo <> WS_SALDO Then
'   Screen.MousePointer = 0
'   MsgBox "..", 48, Pub_Titulo
 '  If LK_CODUSU = "ADMIN" Then GoTo salta
'   pu_flag = 3
'End If
SALTA:

'If pu_ultimo = True Then
'  pu_can = GRIDG.Rows - Val(FrmCaa.txtnum.text)
'  If pu_can < 1 Then
'     pu_can = 1
'  End If
'End If
sigue:
Screen.MousePointer = 0
Exit Sub

Impri:
  If cIngreso = 0 And cSalida = 0 Then GoTo pasecito
  
  fila = fila + 1
  FrmCaa.GRIDG.Rows = FrmCaa.GRIDG.Rows + 1
  GRIDG.RowHeight(fila) = GRIDG.RowHeight(fila) + 50
  GRIDG.TextMatrix(fila, 0) = cFecha
  GRIDG.TextMatrix(fila, 1) = cFecha_emi
  If wcar_mayor!CAA_ESTADO = "E" Then
     GRIDG.Row = fila
     GRIDG.COL = 1
'     GRIDG.CellFontBold = True
     GRIDG.CellBackColor = vb3DLight
     GRIDG.COL = 2
     GRIDG.CellBackColor = vb3DLight
     GRIDG.COL = 3
     GRIDG.CellBackColor = vb3DLight
     GRIDG.COL = 4
     GRIDG.CellBackColor = vb3DLight
     GRIDG.COL = 5
     GRIDG.CellBackColor = vb3DLight
     GRIDG.COL = 6
     GRIDG.CellBackColor = vb3DLight
     GRIDG.COL = 7
     GRIDG.CellBackColor = vb3DLight
   End If

  
  GRIDG.TextMatrix(fila, 2) = cTipdoc
  GRIDG.TextMatrix(fila, 3) = vdocum
  
  GRIDG.TextMatrix(fila, 4) = vCONCEPTO
  If cIngreso <> 0 Then GRIDG.TextMatrix(fila, 5) = Format(cIngreso, "0.00")
  If cSalida <> 0 Then GRIDG.TextMatrix(fila, 6) = Format(cSalida, "0.00")
  GRIDG.TextMatrix(fila, 7) = Format(cSaldo, "0.00")
  GRIDG.TextMatrix(fila, 8) = Format(cFechaV, "dd/mm/yy")
  If cvendedor <> 0 Then GRIDG.TextMatrix(fila, 9) = cvendedor
  GRIDG.TextMatrix(fila, 10) = PUB_CODUSU
  GRIDG.TextMatrix(fila, 11) = Format(chora, "hh:mm:ss")
  GRIDG.TextMatrix(fila, 12) = PUB_FBG
  GRIDG.TextMatrix(fila, 13) = PUB_NUMSER
  GRIDG.TextMatrix(fila, 14) = PUB_NUMFAC
  GRIDG.TextMatrix(fila, 15) = PUB_CODBAN
  If pu_maximo = 99 Then
   GRIDG.TextMatrix(fila, 8) = Format(cFechaV, "dd/mm/yy")
  End If
pasecito:
Return

ACU:
    cIngreso = 0
    cSalida = 0
    If wcar_mayor!CAA_IMPORTE > 0 Then
       cIngreso = cIngreso + wcar_mayor!CAA_IMPORTE
    Else
       cSalida = cSalida + wcar_mayor!CAA_IMPORTE
    End If
    
    If wcar_mayor!CAA_CP = "C" Then
     If Trim(wcar_mayor!CAA_FBG) = "F" Then
        vdocum = "Fact. " & wcar_mayor!CAa_numser & "-" & wcar_mayor!CAa_numfac
     ElseIf Trim(wcar_mayor!CAA_FBG) = "B" Then
        vdocum = "Bolet. " & wcar_mayor!CAa_numser & "-" & wcar_mayor!CAa_numfac
     ElseIf Trim(wcar_mayor!CAA_FBG) = "G" Then
        vdocum = "Guia . " & wcar_mayor!CAa_numser & "-" & wcar_mayor!CAa_numfac
     ElseIf Trim(wcar_mayor!CAA_FBG) = "N" Then
        vdocum = "N.Cred. " & wcar_mayor!CAa_numser & "-" & wcar_mayor!CAa_numfac
     ElseIf Trim(wcar_mayor!CAA_FBG) = "D" Then
        vdocum = "N.Deb. " & wcar_mayor!CAa_numser & "-" & wcar_mayor!CAa_numfac
     Else
       vdocum = "Guia . " & wcar_mayor!CAa_numser & "-" & wcar_mayor!CAa_numfac
     End If
     
     End If
     If wcar_mayor!CAA_TIPMOV = 97 And wcar_mayor!CAA_FBG <> "N" Then
        vdocum = "N.-Cred. " & wcar_mayor!caa_numser_c & "-" & wcar_mayor!caa_numfac_c & " Doc.: " & wcar_mayor!CAA_FBG & " /" & wcar_mayor!CAa_numser & "-" & wcar_mayor!CAa_numfac
     End If
     If wcar_mayor!CAA_TIPMOV = 97 And wcar_mayor!CAA_FBG = "N" Then
        vdocum = "N.-Cred. " & wcar_mayor!CAa_numser & "-" & wcar_mayor!CAa_numfac
     End If
     If wcar_mayor!CAA_TIPMOV = 98 And wcar_mayor!CAA_FBG <> "N" Then
        vdocum = "N.-Deb. " & wcar_mayor!caa_numser_c & "-" & wcar_mayor!caa_numfac_c & " Doc.: " & wcar_mayor!CAA_FBG & " /" & wcar_mayor!CAa_numser & "-" & wcar_mayor!CAa_numfac
     End If
     If wcar_mayor!CAA_TIPMOV = 98 And wcar_mayor!CAA_FBG = "D" Then
        vdocum = "N.-Deb. " & wcar_mayor!CAa_numser & "-" & wcar_mayor!CAa_numfac
     End If
     If wcar_mayor!CAA_CP = "P" And wcar_mayor!CAA_TIPMOV = 98 And wcar_mayor!CAA_FBG = "A" Then
        vdocum = "N.-Deb. " & wcar_mayor!caa_numser_c & "-" & wcar_mayor!caa_numfac_c
     End If
    'verificar porque se puso abelardo...
  '  If wcar_mayor!CAA_TIPDOC <> "NA" And wcar_mayor!CAA_TIPDOC <> "ND" And wcar_mayor!CAA_TIPDOC <> "NC" And wcar_mayor!CAA_TIPDOC <> "FA" And wcar_mayor!CAA_TIPDOC <> "CC" Then
  '     vdocum = ""
  '  End If
    If wcar_mayor!CAA_TIPDOC = "CH" Then
       vdocum = Nulo_Valors(wcar_mayor!CAa_NUM_CHEQUE)
    End If
    If wcar_mayor!CAA_CP = "P" Then
       If wcar_mayor!CAA_TIPMOV = 97 Or wcar_mayor!CAA_TIPMOV = 98 Then
       Else
          If val(Nulo_Valor0(wcar_mayor!caa_numfac_c)) > 0 Then
             vdocum = "Fact. " & wcar_mayor!caa_numser_c & "-" & wcar_mayor!caa_numfac_c
          Else
             vdocum = "Guia . " & Trim(wcar_mayor!CAa_numGUIA)
          End If
       End If
    End If
    
    If wcar_mayor!CAA_TIPDOC = "LE" Then
       vdocum = "Letra."
    End If
    vCONCEPTO = Trim(wcar_mayor!caa_concepto)
    If Left(Trim(wcar_mayor!caa_concepto), 1) >= "0" And Left(Trim(wcar_mayor!caa_concepto), 1) <= "9" Then
       vCONCEPTO = " " + Trim(wcar_mayor!caa_concepto)
    End If
    WS_SALDO = Nulo_Valor0(wcar_mayor!CAA_SALDO)
    cFecha = Format(wcar_mayor!CAA_FECHA, "dd/mm/yy")
    cTipdoc = wcar_mayor!CAA_TIPDOC
    WS_SALDO_ULT = wcar_mayor!CAA_SALDO
    If LK_EMP = "PAR" Then
     cFechaV = wcar_mayor!CAA_FECHA_COBRO
    Else
     cFechaV = wcar_mayor!CAA_FECHA_VCTO
    End If
    cvendedor = Nulo_Valor0(wcar_mayor!CAA_CODVEN)
    PUB_FBG = Nulo_Valors(wcar_mayor!CAA_FBG)
    PUB_NUMFAC = Nulo_Valor0(wcar_mayor!CAa_numfac)
    PUB_NUMSER = Nulo_Valor0(wcar_mayor!CAa_numser)
    PUB_CODUSU = Nulo_Valors(wcar_mayor!CAA_CODUSU)
    chora = DatePart("h", wcar_mayor!CAA_hora) & ":" & DatePart("n", wcar_mayor!CAA_hora)
 '   If LK_EMP = "PLA" Then vdocum = Nulo_Valor0(wcar_mayor!CAA_RECIBO)
    If Nulo_Valor0(wcar_mayor!CAA_RECIBO) > 0 Then vCONCEPTO = "R." & Nulo_Valor0(wcar_mayor!CAA_RECIBO) & "-" & vCONCEPTO
 
    PUB_CODBAN = Nulo_Valor0(wcar_mayor!caa_codban)
     
Return

End Sub

Private Sub ultimo_Click()
If ultimo.Value = 1 Then
  txtnum.Text = "5"
  CmdProcesa.SetFocus
Else
  txtnum.Text = ""
  CmdProcesa.SetFocus
End If

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
wPSCAR_MAYOR(2) = val(i_codcli.Text)
wPSCAR_MAYOR(3) = #1/1/1998# 'pu_fecha
wPSCAR_MAYOR(4) = #1/1/1999# 'pu_fecha
wcar_mayor.Requery
pu_flag = 0
If wcar_mayor.EOF Then
   Screen.MousePointer = 0
   pu_flag = 1
   GoTo sigue
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

sigue:
End Sub
