VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_cheques 
   Caption         =   "Consulta de Bancos"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   1350
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "frm_cheques.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   9510
   WindowState     =   2  'Maximized
   Begin ComctlLib.ListView LV_CCM 
      Height          =   375
      Left            =   7560
      TabIndex        =   12
      Tag             =   "0"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
   Begin VB.CommandButton salir 
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
      Height          =   525
      Left            =   9840
      Picture         =   "frm_cheques.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5880
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   0
      TabIndex        =   15
      Top             =   600
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Resumen "
      TabPicture(0)   =   "frm_cheques.frx":058C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Lche(0)"
      Tab(0).Control(1)=   "Lche(1)"
      Tab(0).Control(2)=   "LMONEDA"
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(4)=   "tipdoc"
      Tab(0).Control(5)=   "txtnro"
      Tab(0).Control(6)=   "grdr"
      Tab(0).Control(7)=   "txt_monto"
      Tab(0).Control(8)=   "txtsig"
      Tab(0).Control(9)=   "txtant"
      Tab(0).Control(10)=   "OPBANCO(0)"
      Tab(0).Control(11)=   "OPBANCO(1)"
      Tab(0).Control(12)=   "resumen"
      Tab(0).Control(13)=   "Command1"
      Tab(0).Control(14)=   "OPBANCO(2)"
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Kardex "
      TabPicture(1)   =   "frm_cheques.frx":05A8
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label3(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtfecha"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "grid_che"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "SaldoLiq"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Saldocon"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "FLAG_V"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdimp"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Ver"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      Begin VB.OptionButton OPBANCO 
         Caption         =   "Transferencias"
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
         Index           =   2
         Left            =   -74760
         TabIndex        =   3
         Top             =   840
         Width           =   1695
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
         Height          =   525
         Left            =   -72480
         Picture         =   "frm_cheques.frx":05C4
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   5280
         Width           =   1455
      End
      Begin VB.CommandButton resumen 
         Caption         =   "R&esumen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -74640
         Picture         =   "frm_cheques.frx":08CE
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   5280
         Width           =   1455
      End
      Begin VB.OptionButton OPBANCO 
         Caption         =   "Abonos"
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
         Index           =   1
         Left            =   -74760
         TabIndex        =   2
         Top             =   560
         Width           =   1215
      End
      Begin VB.OptionButton OPBANCO 
         Caption         =   "Cargos"
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
         Index           =   0
         Left            =   -74760
         TabIndex        =   1
         Top             =   315
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton txtant 
         Caption         =   "<&Ant."
         Height          =   375
         Left            =   -69120
         TabIndex        =   6
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton txtsig 
         Caption         =   "&Sig>"
         Height          =   375
         Left            =   -68400
         TabIndex        =   7
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txt_monto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -66240
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   480
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid grdr 
         Height          =   4095
         Left            =   -74760
         TabIndex        =   25
         Top             =   1080
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7223
         _Version        =   393216
         BackColorBkg    =   9128212
         AllowUserResizing=   1
      End
      Begin VB.CommandButton Ver 
         Caption         =   "&Mostrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   360
         Picture         =   "frm_cheques.frx":0D10
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5280
         Width           =   1455
      End
      Begin VB.CommandButton cmdimp 
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2520
         Picture         =   "frm_cheques.frx":1152
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5280
         Width           =   1455
      End
      Begin VB.TextBox txtnro 
         Height          =   285
         Left            =   -70680
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox tipdoc 
         Height          =   315
         ItemData        =   "frm_cheques.frx":145C
         Left            =   -72960
         List            =   "frm_cheques.frx":1463
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox FLAG_V 
         Caption         =   "Mostrar Extornos"
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
         Left            =   240
         TabIndex        =   21
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Saldocon 
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
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox SaldoLiq 
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
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid grid_che 
         Height          =   4335
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   7646
         _Version        =   393216
         Rows            =   13
         Cols            =   16
         BackColorBkg    =   9128212
         AllowUserResizing=   3
         FormatString    =   $"frm_cheques.frx":1472
      End
      Begin MSMask.MaskEdBox txtfecha 
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "Monto:"
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
         Left            =   -67440
         TabIndex        =   28
         Top             =   480
         Width           =   615
      End
      Begin VB.Label LMONEDA 
         Alignment       =   1  'Right Justify
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
         Left            =   -66840
         TabIndex        =   27
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Lche 
         Caption         =   "Nro.:"
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
         Left            =   -71160
         TabIndex        =   24
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Lche 
         Caption         =   "Tipdoc."
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
         Index           =   0
         Left            =   -72960
         TabIndex        =   23
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "A Partir de :"
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
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Saldo Liquido :"
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
         Index           =   0
         Left            =   7320
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Saldo Contable :"
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
         Left            =   4200
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.TextBox i_codban 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Text            =   "i_codban"
      Top             =   120
      Width           =   1455
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   0
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   4200
      TabIndex        =   31
      Top             =   6480
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Label lblbarraos 
      Alignment       =   2  'Center
      BackColor       =   &H008B4914&
      Caption         =   "Solution for Business"
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
      Left            =   0
      TabIndex        =   33
      Top             =   6840
      Width           =   11895
   End
   Begin VB.Label lblmensaje 
      BackStyle       =   0  'Transparent
      Caption         =   "Procesando . ."
      Height          =   255
      Left            =   3000
      TabIndex        =   32
      Top             =   6480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label i_nomban 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   3000
      TabIndex        =   13
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo de Banco:"
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
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frm_cheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ws_signo_ssm As Integer
Dim xl As Object
Dim PS_CHE As rdoQuery
Dim cheest_llave As rdoResultset
Dim res_cheque As rdoResultset
Dim PSRES_CHE As rdoQuery
Dim loc_key As Integer


Private Sub cmdimp_Click()
Dim i, J
Dim wranF, wran1, wran2
Dim LETRAS(24) As String * 1
If grid_che.Rows < 1 Or grid_che.Visible = False Then
 MsgBox "Nuestre una Consulta.", 48, Pub_Titulo
  Exit Sub
End If
Dim RR As Excel.Application
Dim xl As Object
'On Error GoTo FINTODO
Screen.MousePointer = 11
GoSub WEXCEL
pub_cadena = ""

xl.Cells(4, 1) = "Cta Cte. : " & Trim(i_codban.Text) & " - " & Trim(i_nomban.Caption)
'xl.Cells(5, 1) = "FECHA INICIAL : " & Trim(i_fecha.Text)
xl.Cells(3, 1) = "'" & Format(LK_FECHA_DIA, "dd/mm/yyyy")
xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
xl.Cells(2, 1) = "ESTADO DE CUENTA"
xl.Cells(3, 1) = "'" & Format(LK_FECHA_DIA, "dd/mm/yyyy")
For i = 0 To grid_che.Rows - 1
  For J = 0 To 13
     If grid_che.TextMatrix(i, J) = "" Then
       xl.Cells(i + 7, J + 1) = " "
     Else
       If J = 0 And i > 0 Then
         xl.Cells(i + 7, J + 1) = Trim(grid_che.TextMatrix(i, J))
       Else
        If IsDate(grid_che.TextMatrix(i, J)) Then
            xl.Cells(i + 7, J + 1) = "'" & CStr(grid_che.TextMatrix(i, J))
        Else
         xl.Cells(i + 7, J + 1) = CStr(grid_che.TextMatrix(i, J))
        End If
       End If
     End If
  Next J
  'If i <> 0 Then
  If Trim(grid_che.TextMatrix(i, 12)) = "T" Then
       'xl.Cells(i, J).Cells.Font.Background = QBColor(12)
       xl.Worksheets("Hoja1").Cells(i + 7, 9).Interior.ColorIndex = 6
   End If
  'End If
  'xl.Application.Visible = True
'  If i = 1 Then xl.Cells(i + 7, 4) = "Saldo Incial: "
Next i
GoSub LETRAS

For J = 8 To 9
  wranF = Trim(LETRAS(J)) & i + 1 + 6
  wran1 = Trim(LETRAS(J)) & 8
  wran2 = Trim(LETRAS(J)) & i + 6
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
Next
wranF = "H" & i + 7 & ":" & "L" & i + 7
xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 3

wranF = "A8:" & "L8"
xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 3
xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
xl.Cells(2, 1) = "ESTADO DE CUENTA"
xl.Cells(3, 1) = "'" & Format(LK_FECHA_DIA, "dd/mm/yyyy")
xl.Cells(5, 8) = "CONTABLE:"
xl.Cells(5, 10) = Saldocon.Text
xl.Cells(5, 11) = "LIQUIDO:"
xl.Cells(5, 12) = SaldoLiq.Text

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
  xl.Workbooks.Open Left(Trim(PUB_RUTA_OTRO), 1) & ":\ADMIN\STANDAR\CTABAN.xls", 0, True, 4, PUB_CLAVE, PUB_CLAVE
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

FINTODO:
 'MsgBox ERR"Reintente Nuevamente ..", 48, Pub_Titulo
 MsgBox "Reintente Nuevamente ..", 48, Pub_Titulo

End Sub

Private Sub Command1_Click()
Dim WCODCIA1  As String
Dim WCODCIA2  As String

Reportes.Connect = PUB_ODBC
Reportes.Destination = crptToWindow
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390
If Val(Nulo_Valors(par_llave!par_art_cias)) <> 0 Then
 Reportes.Formulas(1) = "CIA='" & Trim(GEN!GEN_NOMBRE) & "'"
Else
 Reportes.Formulas(1) = "CIA='" & Trim(par_llave!PAR_NOMBRE) & "'"
End If
If Trim(Nulo_Valors(par_llave!par_art_cias)) <> "" Then
 WCODCIA1 = "01"
 WCODCIA2 = "02"
 PSRES_CHE(0) = "01"
 PSRES_CHE(1) = "02"
Else
 WCODCIA1 = LK_CODCIA
End If
PU_NUMFAC = Val((txtnro.Text))
Reportes.WindowTitle = "KARDEX Nº :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
Reportes.ReportFileName = PUB_RUTA_OTRO + "VOCCM.RPT"
pub_cadena = "({ALLOG.ALL_CODCIA} = '" & WCODCIA1 & "' OR {ALLOG.ALL_CODCIA} = '" & WCODCIA2 & "' ) AND {ALLOG.ALL_CODBAN}= " & Trim(i_codban.Text) & "  AND {ALLOG.ALL_CHENUM} = " & PU_NUMFAC & "  AND {ALLOG.ALL_SIGNO_CCM} =" & ws_signo_ssm & " AND {CCMAEST.CCM_CODCIA} = '" & par_llave!PAR_CIACCM & "'"
Reportes.SelectionFormula = pub_cadena
Reportes.WindowTitle = Reportes.WindowTitle & " Archivo: " & Trim(Reportes.ReportFileName)
On Error GoTo VERI
Reportes.Action = 1
Exit Sub
VERI:
MsgBox Err.Description, vbInformation, Pub_Titulo

End Sub

Private Sub FLAG_V_Click()
Ver.SetFocus
End Sub

Private Sub Form_Activate()
TIPDOC.ListIndex = 0
End Sub

Private Sub Form_Load()
ws_signo_ssm = -1
pub_cadena = "SELECT * FROM CHEQUES  WHERE CHE_CODCIA = ? AND CHE_CODBAN = ? AND CHE_NUMOPER = ? "
Set PS_CHE = CN.CreateQuery("", pub_cadena)
PS_CHE(0) = 0
PS_CHE(1) = 0
PS_CHE(2) = 0
'PS_CHE(3) = 0
Set cheest_llave = PS_CHE.OpenResultset(rdOpenKeyset, rdConcurValues)
    
    
pub_cadena = "SELECT all_IMPORTE_AMORT, all_fecha_pro, all_fecha_dia, all_signo_ccm, ALL_CTAG1, all_autocon, all_concepto, all_codusu , ALL_NUMFAC, ALL_numser_c , ALL_NUMFAC_C, ALL_CP , ALL_CODCLIE , ALL_MONEDA_CCM, ALL_CODTRA, ALL_NUMOPER, ALL_FECHA_CAN,ALL_FECHA_SUNAT,ALL_CODBAN, ALL_CHENUM, ALL_CODCIA, ALL_IMPORTE  FROM ALLOG WHERE (ALL_CODCIA = ? OR ALL_CODCIA = ?) AND ALL_CODBAN = ? AND ALL_CHENUM = ? AND ALL_FLAG_EXT <> 'E' AND ALL_SIGNO_CCM <> 0 AND (ALL_CODTRA = 2748 OR ALL_CODTRA = 2735 OR ALL_CODTRA = 5714 OR ALL_CODTRA = 2738 OR ALL_CODTRA = 2720 OR ALL_CODTRA = 5318 )"
Set PSRES_CHE = CN.CreateQuery("", pub_cadena)
PSRES_CHE(0) = 0
PSRES_CHE(1) = 0
PSRES_CHE(2) = 0
PSRES_CHE(3) = 0
Set res_cheque = PSRES_CHE.OpenResultset(rdOpenKeyset, rdConcurValues)

gridr_cabe
grid_che.Cols = 11
txtFecha.Mask = "##/##/##"
txtFecha.Text = Format(DateAdd("d", -30, LK_FECHA_DIA), "dd/mm/yy")
i_nomban.Caption = ""
i_codban.Text = ""
i_codban.MaxLength = 12
End Sub

Private Sub grdr_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Azul txtnro, txtnro
End Sub

Private Sub grid_che_DblClick()
Dim wfecha
If grid_che.COL = 12 Then
    PS_CHE(0) = LK_CODCIA
    PS_CHE(1) = grid_che.TextMatrix(grid_che.Row, 1)
    PS_CHE(2) = Val(i_codban.Text)
    PS_CHE(3) = grid_che.TextMatrix(grid_che.Row, 14)
    cheest_llave.Requery
    If cheest_llave.EOF Then
       MsgBox "Intente Nuevamnete...", 48, Pub_Titulo
       Exit Sub
    End If
    wfecha = InputBox("Cambiar Fecha Para Cobro ..Actual : " & grid_che.TextMatrix(grid_che.Row, 12), "Cambiar")
    If Not IsDate(wfecha) Then Exit Sub
    cheest_llave.Edit
    cheest_llave!CHE_FECHA_COBRO = wfecha
    If cheest_llave!CHE_FECHA_EMISION <> cheest_llave!CHE_FECHA_COBRO Then
     cheest_llave!che_estado = "T"
    Else
     cheest_llave!che_estado = " "
    End If
    cheest_llave.Update
    Ver_Click
End If

If grid_che.COL = 2 Then
    PS_CHE(0) = LK_CODCIA
    PS_CHE(1) = grid_che.TextMatrix(grid_che.Row, 1)
    PS_CHE(2) = Val(i_codban.Text)
    PS_CHE(3) = grid_che.TextMatrix(grid_che.Row, 14)
    cheest_llave.Requery
    If cheest_llave.EOF Then
       MsgBox "Intente Nuevamnete...", 48, Pub_Titulo
       Exit Sub
    End If
    wfecha = InputBox("Cambiar Fecha Emisión : " & grid_che.TextMatrix(grid_che.Row, 2), "Cambiar")
    If Not IsDate(wfecha) Then Exit Sub
    cheest_llave.Edit
    cheest_llave!CHE_FECHA_EMISION = wfecha
    If cheest_llave!CHE_FECHA_EMISION <> cheest_llave!CHE_FECHA_COBRO Then
     cheest_llave!che_estado = "T"
    Else
     cheest_llave!che_estado = " "
    End If
    cheest_llave.Update
    Ver_Click
End If

If grid_che.COL <> 8 Then Exit Sub
If grid_che.TextMatrix(grid_che.Row, 13) = "T" Or grid_che.TextMatrix(grid_che.Row, 13) = "C" Then
    PS_CHE(0) = LK_CODCIA
    PS_CHE(1) = grid_che.TextMatrix(grid_che.Row, 1)
    PS_CHE(2) = Val(i_codban.Text)
    PS_CHE(3) = grid_che.TextMatrix(grid_che.Row, 14)
    cheest_llave.Requery
    If cheest_llave.EOF Then
       MsgBox "Intente Nuevamnete...", 48, Pub_Titulo
       Exit Sub
    End If
    If grid_che.TextMatrix(grid_che.Row, 13) = "T" Then
      pub_mensaje = "Cobrar el Importe !!! ...   ¿Desea Continuar... ?"
    Else
      pub_mensaje = "Retornar a Importe Defirido!!! ...   ¿Desea Continuar... ?"
    End If
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
       Exit Sub
    End If
    cheest_llave.Edit
    If grid_che.TextMatrix(grid_che.Row, 13) = "T" Then
      cheest_llave!che_estado = "C"
    Else
      cheest_llave!che_estado = "T"
    End If
    cheest_llave.Update
    MsgBox "Monto Cobrado ..", 48, Pub_Titulo
    Ver_Click
End If
End Sub

Private Sub grid_che_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode <> 107 And KeyCode <> 109 Then Exit Sub
Dim WVARMAS As Integer
Dim WVARMENOS As Integer
Dim WROW As Integer
Dim WCOL As Integer
If grid_che.COL = 9 Then
    WROW = grid_che.Row
    WCOL = grid_che.COL
    PS_CHE(0) = LK_CODCIA
    PS_CHE(1) = grid_che.TextMatrix(grid_che.Row, 1)
    PS_CHE(2) = Val(i_codban.Text)
    PS_CHE(3) = grid_che.TextMatrix(grid_che.Row, 14)
    cheest_llave.Requery
    If cheest_llave.EOF Then
       MsgBox "Intente Nuevamnete...", 48, Pub_Titulo
       Exit Sub
    End If
    'wfecha = InputBox("Cambiar Fecha Emisión : " & grid_che.TextMatrix(grid_che.Row, 2), "Cambiar")
    'If Not IsDate(wfecha) Then Exit Sub
    WVARMAS = Val(grid_che.TextMatrix(grid_che.Row - 1, 15))
    If grid_che.Row <> 1 Then WVARMENOS = Val(grid_che.TextMatrix(grid_che.Row + 1, 15))
    
    cheest_llave.Edit
    If KeyCode = 107 Then
      cheest_llave!CHE_NUMOPER2 = WVARMAS + 1
    Else
      cheest_llave!CHE_NUMOPER2 = WVARMENOS - 1
    End If
    'If cheest_llave!CHE_FECHA_EMISION <> cheest_llave!CHE_FECHA_COBRO Then
    ' cheest_llave!che_estado = "T"
    'Else
    ' cheest_llave!che_estado = " "
    'End If
    cheest_llave.Update
    Ver_Click
    On Error GoTo sal
    grid_che.Row = WROW
    grid_che.COL = WCOL
End If
sal:

End Sub

Private Sub grid_che_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then i_codban.SetFocus
'If KeyAscii = 13 Then
' MsgBox grid_che.TextMatrix(grid_che.Row, 9)
'End If
End Sub

Private Sub i_codban_Change()
If i_codban.Text = "" Then
 i_nomban.Caption = ""
 grid_che.Clear
End If
End Sub

Private Sub i_codban_GotFocus()
'grid_che.Visible = False
End Sub

Private Sub i_codban_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not LV_CCM.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And i_codban.Text = "" Then
  loc_key = 1
  Set LV_CCM.SelectedItem = LV_CCM.ListItems(loc_key)
'  LV_CLI.Visible = False
  LV_CCM.ListItems.Item(loc_key).Selected = True
  LV_CCM.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > LV_CCM.ListItems.count Then loc_key = LV_CCM.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > LV_CCM.ListItems.count Then loc_key = LV_CCM.ListItems.count
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
  LV_CCM.ListItems.Item(loc_key).Selected = True
  LV_CCM.ListItems.Item(loc_key).EnsureVisible
  i_codban.Text = Trim(LV_CCM.ListItems.Item(loc_key).Text) & " "
  DoEvents
  i_codban.SelStart = Len(i_codban.Text)
  DoEvents
fin:

End Sub

Private Sub i_codban_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As ListItem    ' Variable FoundItem.

If KeyAscii = 27 Then
   i_codban.Text = ""
   LV_CCM.Visible = False
   Exit Sub
End If

If KeyAscii <> 13 Then
   GoTo fin
End If
On Error GoTo mucho
PUB_CODBAN = Val(i_codban.Text)
If Len(i_codban.Text) = 0 Then
   Exit Sub
End If
On Error GoTo 0
If PUB_CODBAN <> 0 Then
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_CCM_LLAVE
   If ccm_llave.EOF Then
    Azul i_codban, i_codban
    MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
    i_codban.SetFocus
    GoTo fin
   Else
    i_nomban.Caption = ccm_llave(2)
   End If
Else
   If loc_key > LV_CCM.ListItems.count Or loc_key = 0 Then
      Exit Sub
   End If

   valor = UCase(LV_CCM.ListItems.Item(loc_key).Text)
   If Trim(UCase(i_codban.Text)) = Left(valor, Len(Trim(i_codban.Text))) Then
   Else
      Exit Sub
   End If
   
   i_codban.Text = Trim(LV_CCM.ListItems.Item(loc_key).SubItems(1))
   PUB_CODBAN = Val(i_codban.Text)
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_CCM_LLAVE
   i_nomban.Caption = ccm_llave(2)
   If TIPDOC.Visible Then
    OPBANCO(0).SetFocus
   ElseIf Ver.Visible Then
    Ver.SetFocus
   End If
   
End If
LV_CCM.Visible = False
If TIPDOC.Visible Then
 OPBANCO(0).SetFocus
ElseIf Ver.Visible Then
 Ver.SetFocus
End If

fin:
mucho:
End Sub

Private Sub i_codban_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
If Len(i_codban.Text) = 0 Or IsNumeric(i_codban.Text) = True Then
   LV_CCM.Visible = False
   Exit Sub
End If
If LV_CCM.Visible = False And KeyCode <> 13 Or Len(i_codban.Text) = 1 Then
    var = Asc(i_codban.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    numarchi = 9
    If LK_EMP = "3AA" Then
    pu_codcia = par_llave!PAR_CIACCM
    archi = "SELECT * FROM CCMAEST WHERE CCM_CODCIA = '" & pu_codcia & "' AND CCM_NOMBRE BETWEEN '" & i_codban.Text & "' AND  '" & var & "' ORDER BY CCM_NOMBRE"
    Else
    archi = "SELECT * FROM CCMAEST WHERE CCM_CODCIA = '" & LK_CODCIA & "' AND CCM_NOMBRE BETWEEN '" & i_codban.Text & "' AND  '" & var & "' ORDER BY CCM_NOMBRE"
    End If
    PROC_LISVIEW LV_CCM
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If LV_CCM.Visible Then
  Set itmFound = LV_CCM.FindItem(LTrim(i_codban.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > LV_CCM.ListItems.count Then
      LV_CCM.ListItems.Item(LV_CCM.ListItems.count).EnsureVisible
   Else
     LV_CCM.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If



End Sub

Private Sub LV_CCM_DblClick()
 loc_key = LV_CCM.SelectedItem.Index
 i_codban.Text = Trim(LV_CCM.ListItems.Item(loc_key).Text) & " "
 i_codban_KeyPress 13

End Sub

Private Sub LV_CCM_GotFocus()
If loc_key <> 0 Then
 Set LV_CCM.SelectedItem = LV_CCM.ListItems(loc_key)
 LV_CCM.ListItems.Item(loc_key).Selected = True
 LV_CCM.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub LV_CCM_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = LV_CCM.SelectedItem.Index
 i_codban.Text = Trim(LV_CCM.ListItems.Item(loc_key).Text) & " "
End If

End Sub

Private Sub LV_CCM_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 LV_CCM.Visible = False
 i_codban.Text = ""
 i_codban.SetFocus
 Exit Sub
End If
If KeyAscii <> 13 Then
 Exit Sub
End If
LV_CCM_DblClick
End Sub

Private Sub LV_CCM_LostFocus()
LV_CCM.Visible = False
End Sub

Private Sub OPBANCO_Click(Index As Integer)
If Index = 0 Then
 ws_signo_ssm = -1
ElseIf Index = 1 Then
 ws_signo_ssm = 1
ElseIf Index = 2 Then
 ws_signo_ssm = 2
End If
End Sub

Private Sub OPBANCO_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   TIPDOC.SetFocus
End If
End Sub

Private Sub resumen_Click()
Dim wmonto As Currency
Dim HDESCRIP As String
Dim HCUENTA As String
PUB_CODBAN = Val(i_codban.Text)
pu_codcia = LK_CODCIA
SQ_OPER = 1
LEER_CCM_LLAVE
If ccm_llave.EOF Then
   Screen.MousePointer = 0
   MsgBox "Banco no existe...", 48, Pub_Titulo
   Exit Sub
End If
gridr_cabe
If Trim(Nulo_Valors(par_llave!par_art_cias)) <> "" Then
 PSRES_CHE(0) = "01"
 PSRES_CHE(1) = "02"
Else
 PSRES_CHE(0) = LK_CODCIA
 PSRES_CHE(1) = 0
End If
PSRES_CHE(2) = PUB_CODBAN
PSRES_CHE(3) = Val(txtnro.Text)
res_cheque.Requery
If res_cheque.EOF Then
gridr_cabe
Azul txtnro, txtnro
Exit Sub
End If
LblMensaje.Visible = True
PB.Min = 0
PB.max = res_cheque.RowCount
PB.Value = 0
PB.Visible = True
fila = 0
grdr.Visible = False
wmonto = 0
Do Until res_cheque.EOF
PB.Value = PB.Value + 1
If ws_signo_ssm = 2 Then
ws_signo_ssm = res_cheque!ALL_SIGNO_CCM
Else
 If res_cheque!ALL_SIGNO_CCM <> ws_signo_ssm Then
    GoTo SIG
 End If
End If
'DoEvents
fila = fila + 1
grdr.Rows = grdr.Rows + 1
HDESCRIP = res_cheque!ALL_autocon
HCUENTA = Nulo_Valors(res_cheque!ALL_CTAG1)
If res_cheque!ALL_CODTRA <> 5714 Then
 pu_cp = res_cheque!ALL_CP
 pu_codclie = res_cheque!ALL_CODCLIE
 pu_codcia = res_cheque!all_CODCIA
 SQ_OPER = 1
 LEER_CLI_LLAVE
 HDESCRIP = Trim(cli_llave!CLI_NOMBRE)
 HCUENTA = Trim(cli_llave!CLI_CUENTA_CONTAB)
End If
grdr.TextMatrix(fila, 0) = Format(res_cheque!ALL_FECHA_CAN, "dd/mm/yy")
grdr.TextMatrix(fila, 1) = HDESCRIP
grdr.TextMatrix(fila, 2) = Format(res_cheque!all_numser_c, "000") & "-" & res_cheque!all_numfac_c
grdr.TextMatrix(fila, 3) = Trim(HCUENTA)
If res_cheque!ALL_CODTRA = 5318 Then
 'If res_cheque!ALL_moneda_ccm = "D" Then
  If res_cheque!ALL_SIGNO_CCM = 1 Then
   grdr.TextMatrix(fila, 4) = Format(Val(res_cheque!ALL_IMPORTE_AMORT), "#,##0.00")
   wmonto = wmonto + Val(res_cheque!ALL_IMPORTE_AMORT)
  Else
   grdr.TextMatrix(fila, 5) = Format(Val(res_cheque!ALL_IMPORTE), "#,##0.00")
   wmonto = wmonto + Val(res_cheque!ALL_IMPORTE)
  End If
Else
 If res_cheque!ALL_moneda_ccm = "D" Then
  grdr.TextMatrix(fila, 4) = Format(Val(res_cheque!ALL_IMPORTE), "#,##0.00")
 Else
  grdr.TextMatrix(fila, 5) = Format(Val(res_cheque!ALL_IMPORTE), "#,##0.00")
 End If
 wmonto = wmonto + Val(res_cheque!ALL_IMPORTE)
End If
grdr.TextMatrix(fila, 6) = Trim(res_cheque!all_concepto)
grdr.TextMatrix(fila, 7) = Trim(res_cheque!all_codusu)
grdr.TextMatrix(fila, 8) = Val(res_cheque!all_numfac)
grdr.TextMatrix(fila, 9) = res_cheque!ALL_CODTRA
grdr.TextMatrix(fila, 10) = res_cheque!all_CODCIA
grdr.TextMatrix(fila, 11) = Format(res_cheque!ALL_FECHA_DIA, "dd/mm/yy")
grdr.TextMatrix(fila, 12) = Format(res_cheque!ALL_FECHA_CAN, "dd/mm/yy")

SIG:
res_cheque.MoveNext
Loop
txt_monto.Text = Format(wmonto, "#,##0.00")
lmoneda.Visible = True
If ccm_llave!CCM_MONEDA = "D" Then
lmoneda.Caption = "US$"
Else
lmoneda.Caption = "S/"
End If
grdr.Visible = True
PB.Visible = False
LblMensaje.Visible = False
If grdr.Rows > 1 Then
grdr.COL = 1
grdr.Row = 1
grdr.SetFocus
End If

End Sub

Private Sub salir_Click()
Unload frm_cheques
End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub tipdoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Azul txtnro, txtnro

End Sub

Private Sub txtant_Click()
If Val(txtnro.Text) <= 0 Then Exit Sub
txtnro.Text = Val(txtnro.Text) - 1
resumen_Click
End Sub

Private Sub txtfecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Ver.SetFocus
End If
End Sub


Private Sub txtnro_Change()
gridr_cabe
End Sub

Private Sub txtnro_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then resumen_Click
If KeyAscii = 27 Then i_codban.SetFocus

SOLO_ENTERO KeyAscii

End Sub

Private Sub txtsig_Click()
'If Val(txtnro.Text) <= 0 Then Exit Sub
txtnro.Text = Val(txtnro.Text) + 1
resumen_Click

End Sub

Private Sub Ver_Click()
Dim fila As Integer
Dim WS_SALDO As Currency
Dim Tit As String
Dim success%
Dim WS_SALDO_DIF As Currency
Dim WS_SALDO_LIQ As Currency
Dim WS_IMPORTE As Currency
Screen.MousePointer = 11
PUB_CODBAN = Val(i_codban.Text)
pu_codcia = LK_CODCIA
SQ_OPER = 1
LEER_CCM_LLAVE
If ccm_llave.EOF Then
   Screen.MousePointer = 0
   MsgBox "Banco no existe...", 48, Pub_Titulo
   GoTo fin
End If
If Not IsDate(txtFecha.Text) Then
  Screen.MousePointer = 0
  MsgBox "Fecha Invalidad..."
  txtFecha.SetFocus
  Exit Sub
End If
fila = 0
PUB_FECHA = txtFecha.Text
SQ_OPER = 4
LEER_CHE_LLAVE
If che_repo.EOF Then
   PUB_FECHA = #1/1/1990#
   LEER_CHE_LLAVE
   If che_repo.EOF Then
      Screen.MousePointer = 0
      MsgBox "No hay estado de cuenta...", 48, Pub_Titulo
      Azul i_codban, i_codban
      GoTo fin
   End If
   
End If
LblMensaje.Visible = True
PB.Min = 0
PB.max = che_repo.RowCount
PB.Visible = True
DoEvents
Screen.MousePointer = 11
che_repo.MoveLast
WS_SALDO = ccm_llave!CCM_SALDO
WS_SALDO_DIF = 0
grid_che.Visible = False
Tit = grid_che.FormatString
grid_che.Clear
grid_che.FormatString = Tit
WS_SALDO_LIQ = ccm_llave!CCM_SALDO
'grid_che.ColWidth(10) = 0
Do Until che_repo.BOF
       PB.Value = fila
       If FLAG_V.Value = 0 And (che_repo!che_estado = "E" Or che_repo!CHE_codtra = 1111) Then GoTo SALT
       fila = fila + 1
       grid_che.Rows = fila + 1
       grid_che.Row = fila
       grid_che.COL = 0
       grid_che.Text = fila & ".-"
       grid_che.COL = 1
       grid_che.Text = Format(che_repo!CHE_FECHA, "dd/mm/yy")
       If che_repo!che_estado = "E" Or che_repo!CHE_codtra = 1111 Then
          grid_che.CellBackColor = vb3DLight
       End If
       grid_che.COL = 2
       grid_che.Text = Format(che_repo!CHE_FECHA_EMISION, "dd/mm/yy")
       grid_che.COL = 3
       grid_che.ColAlignment(2) = 0
       If Left(che_repo!che_cheser, 3) = "i_c" Then
          grid_che.Text = " "
       Else
          grid_che.Text = che_repo!che_cheser
       End If
       If che_repo!che_estado = "E" Or che_repo!CHE_codtra = 1111 Then
          grid_che.CellBackColor = vb3DLight
       End If
       
       grid_che.COL = 4
       grid_che.ColAlignment(3) = 0
       If che_repo!CHE_CHENUM = 0 Then
          grid_che.Text = " "
       Else
          grid_che.Text = che_repo!CHE_CHENUM
       End If
       If che_repo!che_estado = "E" Or che_repo!CHE_codtra = 1111 Then
          grid_che.CellBackColor = vb3DLight
       End If
       
       grid_che.COL = 5
       grid_che.ColAlignment(4) = 0
       If che_repo!CHE_CHENUM = 0 Then
          grid_che.Text = " "
       Else
          grid_che.Text = che_repo!che_chesec
       End If
       If che_repo!che_estado = "E" Or che_repo!CHE_codtra = 1111 Then
          grid_che.CellBackColor = vb3DLight
       End If
       
       grid_che.COL = 6
       grid_che.Text = Nulo_Valors(che_repo!che_abreviado)
       If che_repo!che_estado = "E" Or che_repo!CHE_codtra = 1111 Then
          grid_che.CellBackColor = vb3DLight
       End If
       grid_che.COL = 7
       
       If che_repo!che_estado = "E" Or che_repo!CHE_codtra = 1111 Then
          grid_che.CellBackColor = vb3DLight
       End If
       If che_repo!CHE_SIGNO_CCM = 1 Then
          grid_che.Text = Format(che_repo!CHE_IMPORTE, "0.00")
          grid_che.COL = 8
          grid_che.Text = " "
       Else
          grid_che.Text = " "
          grid_che.COL = 8
          grid_che.Text = Format(che_repo!CHE_IMPORTE, "0.00")
       End If
        If che_repo!che_estado = "E" Or che_repo!CHE_codtra = 1111 Then
           grid_che.CellBackColor = vb3DLight
       End If
       If che_repo!che_estado = "T" Then
         grid_che.CellBackColor = vbYellow
         WS_SALDO_DIF = WS_SALDO_DIF + Val(che_repo!CHE_IMPORTE)
       End If
       grid_che.COL = 9
       grid_che.Text = Format(WS_SALDO, "#,###,##0.00") ' Format(che_repo!che_saldo, "0.00")
       'If WS_SALDO <> che_repo!che_saldo Then
       '   MsgBox "Avisar a Computo...hay diferencia... " & che_repo!che_chenum & che_repo!che_chesec
       'End If
       If WS_SALDO < 0 Then grid_che.CellForeColor = QBColor(12)
       grid_che.COL = 10
       grid_che.Text = Format(WS_SALDO_LIQ, "#,###,##0.00")  ' Format(che_repo!che_saldo, "0.00")
       grid_che.CellForeColor = QBColor(1)
       WS_SALDO = WS_SALDO - che_repo!CHE_SIGNO_CCM * che_repo!CHE_IMPORTE

       If che_repo!che_estado = "E" Or che_repo!CHE_codtra = 1111 Then
       grid_che.CellBackColor = vb3DLight
       End If
       grid_che.COL = 11
       grid_che.ColAlignment(9) = 0
       grid_che.Text = Mid(che_repo!che_concepto, 1, 80)
       If che_repo!che_estado = "E" Or che_repo!CHE_codtra = 1111 Then
          grid_che.CellBackColor = vb3DLight
       End If
       grid_che.COL = 12
       grid_che.Text = Format(che_repo!CHE_FECHA_COBRO, "dd/mm/yy")
       grid_che.COL = 13
       grid_che.Text = Trim(che_repo!che_estado)
       grid_che.COL = 14
       grid_che.Text = che_repo!CHE_NUMOPER
       grid_che.COL = 15
       grid_che.Text = che_repo!CHE_NUMOPER2
SALT:
       che_repo.MovePrevious
Loop

fila = grid_che.Rows - 1
grid_che.Visible = True
WS_SALDO_LIQ = Val(Format(grid_che.TextMatrix(fila, 9), "0.00"))
grid_che.TextMatrix(fila, 10) = Format(WS_SALDO_LIQ, "###,##0.00")
fila = fila - 1
Do Until fila = 0
 If Val(grid_che.TextMatrix(fila, 8)) <> 0 Then
     WS_IMPORTE = Val(grid_che.TextMatrix(fila, 8)) * -1
  Else
     WS_IMPORTE = Val(grid_che.TextMatrix(fila, 7))
  End If
  If grid_che.TextMatrix(fila, 13) <> "T" Then
    WS_SALDO_LIQ = WS_SALDO_LIQ + WS_IMPORTE
  End If
'If fila = grid_che.Rows - 1 Then GoTo PAS
  grid_che.TextMatrix(fila, 10) = Format(WS_SALDO_LIQ, "###,##0.00")
  If WS_SALDO_LIQ < 0 Then
   grid_che.COL = 10
   grid_che.Row = fila
   grid_che.CellForeColor = QBColor(12)
  End If
  
PAS:
fila = fila - 1
Loop
Saldocon.Text = Format(grid_che.TextMatrix(1, 9), "#####0.00")
SaldoLiq.Text = Format(Val(Saldocon.Text) + Val(WS_SALDO_DIF), "####0.00")
If Val(Saldocon.Text) < 0 Then
  Saldocon.ForeColor = QBColor(12)
Else
  Saldocon.ForeColor = QBColor(0)
End If
If Val(SaldoLiq.Text) < 0 Then
  SaldoLiq.ForeColor = QBColor(12)
Else
  SaldoLiq.ForeColor = QBColor(0)
End If
Saldocon.Text = Format(Saldocon.Text, "##,####,##0.00")
SaldoLiq.Text = Format(SaldoLiq.Text, "##,####,##0.00")
LblMensaje.Visible = False
PB.Visible = False
DoEvents
Screen.MousePointer = 0
grid_che.Visible = True
grid_che.Row = 1
grid_che.SetFocus

fin:

End Sub

Private Sub Ver_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then i_codban.SetFocus
End Sub

Public Sub gridr_cabe()
frm_cheques.grdr.Clear
frm_cheques.grdr.Cols = 13
frm_cheques.grdr.Rows = 1
frm_cheques.grdr.TextMatrix(0, 0) = "Fec.Giro"
frm_cheques.grdr.TextMatrix(0, 1) = "Descripción"
frm_cheques.grdr.TextMatrix(0, 2) = "Documento"
frm_cheques.grdr.TextMatrix(0, 3) = "Cuenta"
frm_cheques.grdr.TextMatrix(0, 4) = "Imp.US$"
frm_cheques.grdr.TextMatrix(0, 5) = "Imp.M.N."
frm_cheques.grdr.TextMatrix(0, 6) = "Concepto"
frm_cheques.grdr.TextMatrix(0, 7) = "Cod.Usu"
frm_cheques.grdr.TextMatrix(0, 8) = "Nro.Comp."
frm_cheques.grdr.TextMatrix(0, 9) = "Codtra"
frm_cheques.grdr.TextMatrix(0, 10) = "Alm."
frm_cheques.grdr.TextMatrix(0, 11) = "Fec.Proc"
frm_cheques.grdr.TextMatrix(0, 12) = "Fec.Cont"

frm_cheques.grdr.ColWidth(0) = 800
frm_cheques.grdr.ColWidth(1) = 3300
frm_cheques.grdr.ColWidth(2) = 1200
frm_cheques.grdr.ColWidth(3) = 700
frm_cheques.grdr.ColWidth(4) = 1200
frm_cheques.grdr.ColWidth(5) = 1200
frm_cheques.grdr.ColWidth(6) = 1400
frm_cheques.grdr.ColWidth(7) = 600
frm_cheques.grdr.ColWidth(8) = 600
frm_cheques.grdr.ColWidth(9) = 700
frm_cheques.grdr.ColWidth(10) = 700
frm_cheques.grdr.ColWidth(11) = 800
txt_monto.Text = ""

End Sub
