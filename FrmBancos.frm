VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmBancos 
   Caption         =   "Maestro de Bancos"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11595
   ControlBox      =   0   'False
   Icon            =   "FrmBancos.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   11595
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Agregar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10500
      MaskColor       =   &H00FFFFFF&
      Picture         =   "FrmBancos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1845
      Width           =   1185
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10500
      Picture         =   "FrmBancos.frx":0F2C
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2835
      Width           =   1185
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Modificar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10500
      Picture         =   "FrmBancos.frx":1CEE
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   765
      Width           =   1185
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
      Height          =   735
      Left            =   10500
      Picture         =   "FrmBancos.frx":2B88
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5085
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10500
      Picture         =   "FrmBancos.frx":33FE
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4005
      Width           =   1185
   End
   Begin VB.Frame F1 
      Height          =   1020
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9630
      Begin VB.TextBox txt_key 
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
         Left            =   240
         MaxLength       =   12
         TabIndex        =   21
         Text            =   "cmm_codban"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox TXTBAN 
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
         Index           =   2
         Left            =   2040
         MaxLength       =   40
         TabIndex        =   0
         Text            =   "ccm_nombre"
         Top             =   480
         Width           =   6495
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción :"
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
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo :"
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
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   615
      Left            =   7200
      TabIndex        =   16
      Top             =   5280
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   128
      BackColor       =   14737632
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "pase"
      Height          =   372
      Left            =   8400
      TabIndex        =   20
      Top             =   7440
      Visible         =   0   'False
      Width           =   852
   End
   Begin MSFlexGridLib.MSFlexGrid ListExiste 
      Height          =   735
      Left            =   5280
      TabIndex        =   12
      Top             =   5760
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      _Version        =   393216
      Cols            =   4
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Reporte Bancos"
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   7440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer PARPADEA 
      Interval        =   500
      Left            =   240
      Top             =   5280
   End
   Begin VB.Frame F2 
      Height          =   3255
      Left            =   0
      TabIndex        =   8
      Top             =   1080
      Width           =   9255
      Begin VB.TextBox cta_cont 
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
         Left            =   1680
         MaxLength       =   12
         TabIndex        =   29
         Text            =   "cta_cont"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox TXTBAN 
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
         Index           =   1
         Left            =   1680
         TabIndex        =   25
         Text            =   "cmm_codban"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtalterno 
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
         Left            =   1680
         TabIndex        =   24
         Text            =   "cmm_alterno"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox ccm_nomcorto 
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
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   23
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox moneda 
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
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   240
         WhatsThisHelpID =   18
         Width           =   1215
      End
      Begin ComctlLib.ProgressBar PB 
         Height          =   135
         Left            =   7800
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.CommandButton cmdcontab 
         Caption         =   "Relacionar"
         Height          =   375
         Left            =   7800
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TXTBAN 
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
         Index           =   7
         Left            =   5520
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "cmm_simbolo"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox TXTBAN 
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
         Index           =   4
         Left            =   1680
         MaxLength       =   12
         TabIndex        =   3
         Text            =   "tcontab"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox TXTBAN 
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
         Index           =   3
         Left            =   5520
         TabIndex        =   2
         Text            =   "cmm_saldo"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cuenta  de  Banco.  :"
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
         Height          =   390
         Index           =   8
         Left            =   540
         TabIndex        =   30
         Top             =   2520
         Width           =   1050
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cia.:"
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
         Index           =   3
         Left            =   1185
         TabIndex        =   28
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Codigo Alterno :"
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
         Left            =   210
         TabIndex        =   27
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Abreviado:"
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
         Index           =   7
         Left            =   660
         TabIndex        =   26
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Simbolo :"
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
         Left            =   4440
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda D=Dolares  S=Soles :"
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
         Index           =   5
         Left            =   3000
         TabIndex        =   13
         Top             =   240
         Width           =   3135
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cuenta de Transferencia. :"
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
         Height          =   435
         Index           =   4
         Left            =   225
         TabIndex        =   10
         Top             =   1920
         Width           =   1365
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Saldo Actual :"
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
         Index           =   2
         Left            =   4560
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H008B4914&
      BorderStyle     =   1  'Fixed Single
      Height          =   7095
      Left            =   10320
      TabIndex        =   22
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label momento 
      AutoSize        =   -1  'True
      Caption         =   "Procesando... un momento ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6120
      TabIndex        =   15
      Top             =   7920
      Visible         =   0   'False
      Width           =   3105
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
      Height          =   240
      Left            =   600
      TabIndex        =   11
      Top             =   3600
      Width           =   420
   End
End
Attribute VB_Name = "frmBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim llave1 As String
Dim CU As Integer
Dim pasa As Boolean
Dim loc_key  As Integer
Dim LOC_CANCELA As Integer
Dim LOC_CTA_CLI As String * 12
Dim LOC_DES_CLI As String * 50
Dim LOC_NIVEL As Integer
Dim LOC_CTA_SUP As String
Dim LOC_FLAG_AFEC As String * 1
Dim LOC_ESTADO As String * 1
Dim LOC_TIPO_CTA As Integer
Dim LOC_SIGNO_D As Integer
Dim LOC_SIGNO_H As Integer
Dim LOC_ACT_PAS As Integer

Public Function PROC_BAN()
Dim fila As Integer
Dim NUMCAMPO As Integer
Dim wBusca As String
Dim OJO As String * 1
If allcias.Value = 0 Then
  wBusca = "SELECT * FROM  CCMAEST WHERE  CCM_NOMBRE >= ? AND CCM_CODCIA = ? ORDER BY CCM_NOMBRE"
ElseIf allcias.Value = 1 Then
  wBusca = "SELECT * FROM  CCMAEST WHERE  CCM_NOMBRE >= ?  ORDER BY CCM_NOMBRE"
End If
If UNICO <> wBusca Then
   Set PSX = CN.CreateQuery("", wBusca)
End If
PSX.rdoParameters(0) = LEIDO.Text
If allcias.Value = 0 Then
  If LK_EMP = "3AA" Then
    PSX.rdoParameters(1) = par_llave!PAR_CIACCM
  Else
    PSX.rdoParameters(1) = LK_CODCIA
  End If
End If
If UNICO <> wBusca Then
   Set X = PSX.OpenResultset(rdOpenKeyset)
End If
   
If UNICO = wBusca Then
   X.Requery
   If X.RowCount > 0 Then
      X.MoveFirst
   End If
End If
If X.EOF = True Then
   grid1.Clear
   grid1.Visible = True
   grid1.Row = 1
   grid1.COL = 1
   grid1.Text = "No hay registros"
   llave1 = ""
   Exit Function
End If

If X.rdoColumns(0) = llave1 Then
    Exit Function
End If
If X.RowCount > 0 Then
   llave1 = X.rdoColumns(0)
End If
Screen.MousePointer = 11
UNICO = wBusca
grid1.Rows = 2
fila = 0
grid1.Clear
Cabecera
LEIDO.Locked = False
grid1.Rows = 2
fila = 0
Do Until X.EOF Or fila = 50
    fila = fila + 1
    grid1.TextMatrix(fila, 0) = Nulo_Valors(X.rdoColumns(1))
    grid1.TextMatrix(fila, 1) = Nulo_Valors(X.rdoColumns(2))
    grid1.TextMatrix(fila, 2) = Nulo_Valors(X.rdoColumns(0))
    X.MoveNext
    grid1.Rows = grid1.Rows + 1
WLOOP:
Loop
grid1.TextMatrix(fila + 1, 1) = "                * * *    END    * * * "
Screen.MousePointer = 0
grid1.Row = 1
grid1.COL = 1
grid1.Visible = True
If LEIDO.Enabled And LEIDO.Visible Then
    LEIDO.SetFocus
End If

End Function

Public Sub Cabecera()
grid1.Cols = 3
grid1.ColWidth(0) = 250
grid1.ColWidth(1) = 3000
grid1.ColWidth(2) = 1
grid1.Row = 0
grid1.COL = 1
grid1.Text = " Descripción "
grid1.ColAlignment(1) = 2
End Sub

Public Function EXISTE_BAN(VALOR1 As String) As Boolean
If VALOR1 <> "" Then
   VALOR1 = "*" & VALOR1 & "*"
Else
   VALOR1 = ""
End If

frmBancos.ListExiste.Row = 0
frmBancos.ListExiste.COL = 0
frmBancos.ListExiste.Text = " CODIGO "
frmBancos.ListExiste.COL = 1
frmBancos.ListExiste.Text = " DESCRIPCION "
frmBancos.ListExiste.COL = 2
frmBancos.ListExiste.Text = " CIA. "
frmBancos.ListExiste.COL = 3
frmBancos.ListExiste.Text = " SALDO  "

frmBancos.ListExiste.Left = 600
frmBancos.ListExiste.Top = 2880
frmBancos.ListExiste.Height = 1335
frmBancos.ListExiste.Width = 8055
EXISTE_BAN = False

Dim NUMCAMPO As Integer
Dim OJO As String * 1
Static P As Boolean
Dim var As String
Dim chec1 As Boolean
Dim chec2 As Boolean
Dim chec3 As Boolean
PUB_CODBAN = 0
SQ_OPER = 2
pu_codcia = LK_CODCIA
LEER_CCM_LLAVE
If ccm_mayor.EOF = True Then
   GoTo fin
End If
ccm_mayor.MoveFirst
fila = 0
frmBancos.ListExiste.Rows = 2
Do Until ccm_mayor.EOF 'Or fila = 50
If ccm_mayor!CCM_CODCIA <> LK_CODCIA Then
  GoTo SIG
End If
OJO = "S"
'VAR = "*" & VALOR & "*"
chec1 = Nulo_Valors(ccm_mayor!CCM_NOMBRE) Like VALOR1
If chec1 = False Then
   OJO = "N"
End If

If OJO = "S" Then
    fila = fila + 1
    frmBancos.ListExiste.Row = fila
    frmBancos.ListExiste.COL = 0
    frmBancos.ListExiste.Text = Nulo_Valors(ccm_mayor!CCM_CODBAN)
    frmBancos.ListExiste.COL = 1
    frmBancos.ListExiste.Text = Nulo_Valors(ccm_mayor!CCM_NOMBRE)
    frmBancos.ListExiste.COL = 2
    frmBancos.ListExiste.Text = Nulo_Valors(ccm_mayor!CCM_CODCIA)
    frmBancos.ListExiste.COL = 3
    frmBancos.ListExiste.Text = Nulo_Valors(ccm_mayor!CCM_SALDO)
    EXISTE_BAN = True
    frmBancos.ListExiste.Rows = frmBancos.ListExiste.Rows + 1
End If
SIG:
    ccm_mayor.MoveNext
Loop

If EXISTE_BAN Then
    frmBancos.ListExiste.Row = frmBancos.ListExiste.Row + 1
    frmBancos.ListExiste.COL = 0
    frmBancos.ListExiste.Text = "*** END ***"
    frmBancos.ListExiste.Visible = True
    frmBancos.ListExiste.Row = 1
    frmBancos.ListExiste.COL = 1
    frmBancos.ListExiste.SetFocus

End If

GoTo fin
Exit Function

CHECKERROR:
MsgBox Err.Description
fin:

End Function


Public Function GENERA_BAN() As Double
Dim NUMCAD, FIJO As String
Dim DIGI As String * 2
Dim I, VINT1, VINT2, VINT3, VINT4 As Double
Dim VSTR1, VSTR2, VSTR3, VSTR4 As String
Dim VFIJO As Double
Dim VVARI As Integer
Dim STRpub_cadena As String
Dim INTpub_cadena As Double
PUB_CODBAN = 0
pu_codcia = LK_CODCIA
SQ_OPER = 2
LEER_CCM_LLAVE
If ccm_mayor.EOF Then
    NUMCAD = "207"
Else
    ccm_mayor.MoveLast
    NUMCAD = ccm_mayor!CCM_CODBAN
End If

VINT2 = 0
NUMCAD = Trim(NUMCAD)
VINT1 = Len(NUMCAD)
If NUMCAD = "1" Or NUMCAD = "2" Or NUMCAD = "0" Then
  VINT2 = 1
  VINT1 = 2
End If
If VINT1 > 1 Then
    VSTR4 = Val(Mid(NUMCAD, 1, VINT1 - 2)) + 1
End If

For I = 1 To VINT1 - 2
   VSTR1 = Mid(VSTR4, I, 1)
   VINT2 = VINT2 + Val(VSTR1)
Next I
VINT3 = (VINT2 * 3) + 1

VSTR3 = Right(CStr(VINT3), 2)
If Len(VSTR3) = 1 Then
  VSTR3 = "0" & VSTR3
End If
FIJO = VSTR4
STRpub_cadena = FIJO & VSTR3
INTpub_cadena = Val(STRpub_cadena)
GENERA_BAN = INTpub_cadena

End Function

Public Sub GRABAR_BAN()
If Left(CmdModificar.Caption, 2) = "&G" Then
   ccm_llave.Edit
Else
   ccm_llave.AddNew
       
End If

ccm_llave!CCM_CODBAN = Val(frmBancos.Txt_key)
ccm_llave!CCM_CODCIA = frmBancos.TXTBAN(1)
ccm_llave!CCM_NOMBRE = frmBancos.TXTBAN(2)
ccm_llave!CCM_SALDO = Val(frmBancos.TXTBAN(3))
ccm_llave!CCM_CUENTA_CONTAB = frmBancos.TXTBAN(4)
ccm_llave!CCM_ALTERNO = Trim(txtalterno.Text)
ccm_llave!CCM_MONEDA = " "
If Left(frmBancos.moneda.Text, 1) = "S" Then
  ccm_llave!CCM_MONEDA = "S"
ElseIf Left(frmBancos.moneda.Text, 1) = "D" Then
  ccm_llave!CCM_MONEDA = "D"
End If
ccm_llave!CCM_SIMBOLO = Trim(frmBancos.TXTBAN(7))
ccm_llave!ccm_nomcorto = Trim(ccm_nomcorto.Text)
ccm_llave!ccm_cuenta_contab2 = Trim(frmBancos.cta_cont.Text)
ccm_llave.Update
End Sub
Public Sub MENSAJE_BAN(TEXTO As String)
  LblMensaje.Caption = TEXTO
  PARPADEA.Enabled = True
End Sub

Public Sub LLENA_BAN(ban As Integer)
Dim I As Integer
    If ban = 0 Then
       '**  BAN = 0 BUSCA DATOS NUEVAMENTE
       If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
        Else
          Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
       End If
       SQ_OPER = 1
       PUB_CODBAN = Val(Txt_key.Text)
       pu_codcia = LK_CODCIA
       LEER_CCM_LLAVE
    End If
    
For I = 1 To 3
    frmBancos.TXTBAN(I).Text = Nulo_Valors(Trim(ccm_llave.rdoColumns(I)))
Next I
If Trim(Nulo_Valors(ccm_llave.rdoColumns(6))) = "S" Then
  frmBancos.moneda.ListIndex = 0
ElseIf Trim(Nulo_Valors(ccm_llave.rdoColumns(6))) = "D" Then
  frmBancos.moneda.ListIndex = 1
End If
frmBancos.TXTBAN(7).Text = Trim(Nulo_Valors(ccm_llave.rdoColumns(7)))
frmBancos.TXTBAN(4).Text = Trim(Nulo_Valors(ccm_llave!CCM_CUENTA_CONTAB))
cta_cont.Text = Trim(Nulo_Valors(ccm_llave!ccm_cuenta_contab2))
txtalterno.Text = Trim(Nulo_Valors(ccm_llave!CCM_ALTERNO))
ccm_nomcorto.Text = Trim(Nulo_Valors(ccm_llave!ccm_nomcorto))
If Trim(Nulo_Valors(ccm_llave!CCM_CUENTA_CONTAB)) <> "" Then
    cmdcontab.Caption = "&Quitar"
Else
    cmdcontab.Caption = "Re&lacionar"
End If




End Sub
Public Sub LIMPIA_BAN()
Dim I As Integer
For I = 1 To 4
    TXTBAN(I).Text = ""
Next I
TXTBAN(7).Text = ""
frmBancos.moneda.ListIndex = -1
txtalterno.Text = ""
Txt_key.Text = ""
ccm_nomcorto.Text = ""
cta_cont.Text = ""
End Sub

Private Sub ccm_nomcorto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     moneda.SetFocus
End If
End Sub

Private Sub cmdagregar_Click()
'On Error GoTo ESCAPA
If Left(CmdAgregar.Caption, 2) = "&A" Then
    CmdAgregar.Caption = "&Grabar"
    cmdCancelar.Enabled = True
    CmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    LIMPIA_BAN
    Txt_key.Locked = False
    TXTBAN(2).Enabled = True
    frmBancos.moneda.Enabled = True
    TXTBAN(4).Enabled = True
    TXTBAN(7).Enabled = True
    cmdcontab.Enabled = True
    txtalterno.Enabled = True
    ccm_nomcorto.Enabled = True
    cta_cont.Enabled = True
    frmBancos.Txt_key = GENERA_BAN
    frmBancos.TXTBAN(1) = LK_CODCIA
    frmBancos.TXTBAN(3) = "0.00"
    frmBancos.TXTBAN(4) = ""
    frmBancos.TXTBAN(2).SetFocus
    frmBancos.txtalterno.Text = Trim(frmBancos.Txt_key)
    frmBancos.cmdcontab.Caption = "Re&lacionar"
    'AGREGAMOS EN BLANCO
Else
   If frmBancos.TXTBAN(2).Text = "" Or Len(frmBancos.TXTBAN(2).Text) = 0 Then
       MsgBox "Ingrese Descripción de Banco ..!!!", 48, Pub_Titulo
       Azul TXTBAN(2), TXTBAN(2)
       Exit Sub
   End If
   If Trim(frmBancos.moneda.Text) = "" Then
     MsgBox "Definir la Moneda del Banco. ", 48, Pub_Titulo
     moneda.SetFocus
     Exit Sub
   End If
   '"SI GRABA.."
   If Not pasa Then
     If EXISTE_BAN(Left(frmBancos.TXTBAN(2).Text, 15)) Then
        MENSAJE_BAN " Existen algunos clientes con estos NOMBRES .."
        frmBancos.ListExiste.SetFocus
        Exit Sub
     End If
    End If
    SQ_OPER = 1
    PUB_CODBAN = Val(frmBancos.Txt_key.Text)
    pu_codcia = LK_CODCIA
    LEER_CCM_LLAVE
    If Not ccm_llave.EOF Then
       pasa = False
       frmBancos.ListExiste.Visible = False
       MsgBox "Registro ,  EXISTE ... ", 48, Pub_Titulo
       Azul frmBancos.Txt_key, frmBancos.Txt_key
       Exit Sub
    End If
   frmBancos.ListExiste.Visible = False
   GRABA_CONTAB LK_CODCIA
   Screen.MousePointer = 11
   GRABAR_BAN
   MENSAJE_BAN "Bancos , AGREGADO... "
   CmdAgregar.Caption = "&Agregar"
   CmdModificar.Enabled = True
   cmdEliminar.Enabled = True
   Txt_key.Locked = False
   TXTBAN(2).Enabled = False
   frmBancos.moneda.Enabled = False
   TXTBAN(4).Enabled = False
   TXTBAN(7).Enabled = False
   cmdcontab.Enabled = False
   txtalterno.Enabled = False
   ccm_nomcorto.Enabled = False
   cta_cont.Enabled = False
   LIMPIA_BAN
   Txt_key.SetFocus
   pasa = False
   'BLOQUEA_TEXT
   Screen.MousePointer = 0
      
End If
   
End Sub

Private Sub cmdAgregar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    frmBancos.Txt_key.SetFocus
End If

End Sub

Private Sub CmdCancelar_Click()
If Left(CmdAgregar.Caption, 2) = "&A" And Left(CmdModificar.Caption, 2) = "&M" Then
    LIMPIA_BAN
    Txt_key.Locked = False
    MENSAJE_BAN "Proceso Cancelado... !!!    "
    Txt_key.Enabled = True
    Txt_key.SetFocus
    cmdcontab.Enabled = False
    Exit Sub
End If
     Screen.MousePointer = 11
     If Left(CmdModificar.Caption, 2) = "&G" Then
        CmdModificar.Caption = "&Modificar"
        LLENA_BAN 1
        Txt_key.Locked = True
        TXTBAN(2).Enabled = False
        TXTBAN(4).Enabled = False
        frmBancos.moneda.Enabled = False
        cmdcontab.Enabled = False
        TXTBAN(7).Enabled = False
        txtalterno.Enabled = False
        ccm_nomcorto.Enabled = False
        cta_cont.Enabled = False
     Else
        CmdAgregar.Caption = "&Agregar"
        LIMPIA_BAN
        Txt_key.Locked = False
        TXTBAN(2).Enabled = False
        TXTBAN(4).Enabled = False
        frmBancos.moneda.Enabled = False
        TXTBAN(7).Enabled = False
        txtalterno.Enabled = False
        ccm_nomcorto.Enabled = False
        cta_cont.Enabled = False
     End If
     CmdCerrar.Caption = "Ce&rrar"
     CmdAgregar.Enabled = True
     CmdModificar.Enabled = True
     cmdEliminar.Enabled = True
     cmdcontab.Enabled = False
     Txt_key.Enabled = True
     txtalterno.Enabled = True
     ccm_nomcorto.Enabled = True
     MENSAJE_BAN "Proceso Cancelado... !!!    "
     Screen.MousePointer = 0
End Sub

Private Sub cmdcontab_Click()
Dim wgrupo As String
LK_TABLA = "BANCO"
wgrupo = "10"
archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_CUENTA >= '" & wgrupo & "' AND COM_CUENTA < '" & Trim(Str(Val(wgrupo) + 1)) & "'  ORDER BY COM_CUENTA"
If Left(cmdcontab.Caption, 2) = "&Q" Then
    pub_mensaje = "Confirmar la eliminación de la Cuenta : " & Trim(TXTBAN(4).Text) & " , Continuar ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
       Exit Sub
    End If
    SQ_OPER = 1
    PUB_CUENTA = Trim(TXTBAN(4).Text)
    LEER_COM_LLAVE
    If com_llave.EOF Then
      TXTBAN(4).Text = ""
      CmdModificar_Click
    Else
      com_llave.Delete
      TXTBAN(4).Text = ""
      CmdModificar_Click
    End If
    cmdcontab.Caption = "Relacionar"
    Exit Sub
End If
LOC_CANCELA = 0
LOC_CTA_CLI = ""
LOC_DES_CLI = ""
DoEvents
Load frmBuscacta
frmBuscacta.lbltabla.Caption = LK_TABLA
frmBuscacta.Show 1
LOC_CTA_CLI = Trim(frmBuscacta.tcuenta)
LOC_DES_CLI = Trim(frmBuscacta.tnombre.Text)
LOC_NIVEL = Val(frmBuscacta.txtdatos(0).Text)
LOC_CTA_SUP = Trim(frmBuscacta.txtdatos(1).Text)
LOC_FLAG_AFEC = Trim(frmBuscacta.txtdatos(2).Text)
LOC_ESTADO = Trim(frmBuscacta.txtdatos(3).Text)
LOC_TIPO_CTA = Val(frmBuscacta.txtdatos(4).Text)
LOC_SIGNO_D = Val(frmBuscacta.txtdatos(5).Text)
LOC_SIGNO_H = Val(frmBuscacta.txtdatos(6).Text)
LOC_ACT_PAS = Val(frmBuscacta.txtdatos(7).Text)
TXTBAN(4).Text = Trim(LOC_CTA_CLI)
If Trim(LOC_DES_CLI) = "" And Trim(LOC_DES_CLI) = "" Then
 LOC_CANCELA = 1
Else
 LOC_CANCELA = 0
End If
Unload frmBuscacta

End Sub

Private Sub cmdEliminar_Click()
Dim PS_REP01 As rdoQuery
Dim llave_rep01  As rdoResultset
Dim wcias As String
On Error GoTo SALE
If Len(Txt_key) = 0 Or Len(TXTBAN(1).Text) = 0 Then
   MENSAJE_BAN "NO a seleccionado NADA ... !"
   Exit Sub
End If
  'Dim PS_REP01 As rdoQuery
  'Dim llave_rep01 As rdoResultset
  Screen.MousePointer = 11
  LblMensaje.Visible = True
  LblMensaje.Caption = "Verificando Data.. Un Momento..."
  DoEvents
  pub_cadena = "SELECT CHE_CODBAN FROM CHEQUES WHERE CHE_CODCIA = ? AND CHE_CODBAN = ? "
  Set PS_REP01 = CN.CreateQuery("", pub_cadena)
  PS_REP01.rdoParameters(0) = " "
  PS_REP01.rdoParameters(1) = 0
  PS_REP01.MaxRows = 1
  Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
  PS_REP01(0) = LK_CODCIA
  PS_REP01(1) = ccm_llave!CCM_CODBAN
  llave_rep01.Requery
  LblMensaje.Visible = False
  If Not llave_rep01.EOF Then
     Screen.MousePointer = 0
     MsgBox "NO se Puede Eliminar ...  BANCO  TIENE H I S T O R I A.. ", 48, Pub_Titulo
     Exit Sub
  End If
  Screen.MousePointer = 0
  pub_mensaje = " ¿Desea Eliminar el Registro... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbYes Then   ' El usuario eligió
    Screen.MousePointer = 11
    ccm_llave.Delete
    frmBancos.Txt_key.Text = ""
    frmBancos.Txt_key.Locked = False
    LIMPIA_BAN
    MENSAJE_BAN "Registro   ELIMINADO ... "
    Screen.MousePointer = 0
   Exit Sub
  End If
  Screen.MousePointer = 0
Exit Sub
SALE:
    MsgBox Err.Number & "  " & Err.Description & "  Intente Nuevamente."
    CmdCancelar_Click
    Screen.MousePointer = 0

End Sub

Private Sub cmdcerrar_Click()
ws_conta = 0
Unload frmBancos

End Sub

Private Sub cmdCerrar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    frmBancos.Txt_key.SetFocus
End If

End Sub

Private Sub CmdModificar_Click()
If Len(Txt_key) = 0 Or Len(TXTBAN(2)) = 0 Then
   MENSAJE_BAN "NO a seleccionado NADA ... !"
   Exit Sub
End If
If Left(CmdModificar.Caption, 2) = "&M" Then
    CmdModificar.Caption = "&Grabar"
    CmdAgregar.Enabled = False
    cmdEliminar.Enabled = False
    cmdCancelar.Enabled = True
    Txt_key.Locked = True
    TXTBAN(2).Enabled = True
    TXTBAN(4).Enabled = True
    TXTBAN(7).Enabled = True
    frmBancos.moneda.Enabled = True
    cmdcontab.Enabled = True
    txtalterno.Enabled = True
    ccm_nomcorto.Enabled = True
    cta_cont.Enabled = True
    TXTBAN(2).SetFocus
    TXTBAN(2).SelLength = Len(TXTBAN(2).Text)
    
Else
    '*Grabar las modificaciones
    If TXTBAN(2).Text = "" Or Len(TXTBAN(2).Text) = 0 Then
         MsgBox " Nombre Invalido ....", 48, Pub_Titulo
         Exit Sub
    End If
     Screen.MousePointer = 11
     GRABA_CONTAB LK_CODCIA
     GRABAR_BAN
     MENSAJE_BAN "Registro , MODIFICADO... "
     CmdModificar.Caption = "&Modificar"
     CmdAgregar.Enabled = True
     cmdEliminar.Enabled = True
     Txt_key.Locked = True
     TXTBAN(2).Enabled = False
     TXTBAN(4).Enabled = False
     TXTBAN(7).Enabled = False
     frmBancos.moneda.Enabled = False
     txtalterno.Enabled = False
     ccm_nomcorto.Enabled = False
     cmdcontab.Enabled = False
     cta_cont.Enabled = False
     Screen.MousePointer = 0
End If

End Sub

Private Sub cmdModificar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    frmBancos.Txt_key.SetFocus
End If

End Sub


Private Sub Command2_Click()
Dim WNIVEL_ACT As Integer
Dim PSPRO1 As rdoQuery
Dim pro1_llave As rdoResultset
pub_cadena = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? and COM_CUENTA = ? ORDER BY  COM_CODCIA"
Set PSPRO1 = CN.CreateQuery("", pub_cadena)
PSPRO1(0) = 0
PSPRO1(1) = 0
Set pro1_llave = PSPRO1.OpenResultset(rdOpenKeyset, rdConcurValues)
Data1.Refresh
Print Data1.Recordset.RecordCount
Do Until Data1.Recordset.EOF
  GoSub AGREGA
 Data1.Recordset.MoveNext
Loop
MsgBox "LISTO"
Exit Sub
AGREGA:
    If Trim(Data1.Recordset!MCTAAMARR1) <> "" Then
        PSPRO1(0) = LK_CODCIA
        PSPRO1(1) = Data1.Recordset!MCODCTA
        pro1_llave.Requery
        pro1_llave.Edit
        pro1_llave!com_cuenta_AUTOM_D = Nulo_Valors(Data1.Recordset!MCTAAMARR1)
        pro1_llave!COM_CUENTA_AUTOM_D2 = Nulo_Valors(Data1.Recordset!MCTAAMARR2)
        pro1_llave!COM_CUENTA_AUTOM_D3 = Nulo_Valors(Data1.Recordset!MCTAAMARR3)
        pro1_llave.Update
    End If
Return
   

End Sub

Private Sub Form_Load()

llave1 = ""
CU = 0
pasa = False
LIMPIA_BAN
Txt_key.Enabled = True
TXTBAN(2).Enabled = False
TXTBAN(4).Enabled = False
frmBancos.moneda.Enabled = False
TXTBAN(7).Enabled = False
cmdcontab.Enabled = False
cta_cont.Enabled = False

frmBancos.ListExiste.ColWidth(0) = 900
frmBancos.ListExiste.ColWidth(1) = 3000
frmBancos.ListExiste.ColWidth(2) = 2000
frmBancos.ListExiste.ColWidth(3) = 2000
frmBancos.moneda.Clear
frmBancos.moneda.AddItem "S = Soles"
frmBancos.moneda.AddItem "D = Dolar"
Command1.Enabled = True
SQ_OPER = 1
PUB_CODBAN = 1
pu_codcia = LK_CODCIA
LEER_CCM_LLAVE
If ccm_llave.EOF Then
  crear_caja_chica
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
ws_conta = 0
End Sub

Private Sub ListExiste_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  pub_mensaje = " ¿Desea aun Grabar estos datos de  Bancos ... ? "
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbYes Then   ' El usuario eligió
     pasa = True
     frmBancos.ListExiste.Visible = False
     cmdagregar_Click
     KeyAscii = 0
  Else
    pasa = False
    frmBancos.ListExiste.Visible = False
   ' cmdCerrar_Click
    frmBancos.TXTBAN(2).SetFocus
    KeyAscii = 0
  End If

End If
End Sub

Private Sub ListExiste_LostFocus()
If frmBancos.ListExiste.Visible = False Then
     Exit Sub
End If
  pub_mensaje = " ¿Desea aun Grabar estos datos de  Bancos ... ? "
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbYes Then   ' El usuario eligió
     pasa = True
     cmdagregar_Click
     KeyAscii = 0
  Else
   frmBancos.ListExiste.Visible = False
    'cmdCerrar_Click
   If frmBancos.TXTBAN(2).Enabled Then
       frmBancos.TXTBAN(2).SetFocus
   End If
    KeyAscii = 0
  End If

End Sub

Private Sub ListView1_DblClick()
 loc_key = ListView1.SelectedItem.Index
 Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
 txt_key_KeyPress 13
End Sub

Private Sub ListView1_GotFocus()
If loc_key <> 0 Then
 Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
 ListView1.ListItems.Item(loc_key).Selected = True
 ListView1.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = ListView1.SelectedItem.Index
 Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
End If

End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 ListView1.Visible = False
 Txt_key.Text = ""
 Txt_key.SetFocus
 Exit Sub
End If
If KeyAscii <> 13 Then
 Exit Sub
End If
ListView1_DblClick

End Sub

Private Sub ListView1_LostFocus()
ListView1.Visible = False
End Sub

Private Sub moneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Left(moneda.Text, 1) = "S" Then
   TXTBAN(7).Text = "S/."
  ElseIf Left(moneda.Text, 1) = "D" Then
   TXTBAN(7).Text = "US$"
  End If
  TXTBAN(7).SetFocus
End If
End Sub

Private Sub moneda_LostFocus()
 If Trim(TXTBAN(7).Text) <> "" Then moneda_KeyPress 13
End Sub

Private Sub PARPADEA_Timer()
 CU = CU + 1
 LblMensaje.Visible = True 'Not LblMensaje.Visible
 If CU > 2 Then
   CU = 0
   PARPADEA.Enabled = False
   LblMensaje.Visible = False
 End If

End Sub

Private Sub txt_key_GotFocus()
 Azul Txt_key, Txt_key
End Sub
Private Sub txt_key_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And Txt_key.Text = "" Then
  loc_key = 1
  Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  Txt_key.SelStart = Len(Txt_key.Text)
fin:

End Sub
Private Sub txt_key_KeyPress(KeyAscii As Integer)
Dim VALOR As String
Dim tf As Integer
Dim I
Dim itmFound As ListItem
'On Error GoTo SALCODI
If KeyAscii = 27 And Trim(TXTBAN(1).Text) = "" Then
 Txt_key.Text = ""
End If
If KeyAscii <> 13 Or Left(CmdAgregar.Caption, 2) = "&G" Or Left(CmdModificar.Caption, 2) = "&G" Then
   GoTo fin
End If
pu_codclie = Val(Txt_key.Text)
If Len(Txt_key.Text) = 0 Then
   Exit Sub
End If
'fra2.Refresh
If pu_codclie <> 0 And IsNumeric(Txt_key.Text) = True Then
    SQ_OPER = 1
    On Error GoTo mucho
    PUB_CODBAN = Val(frmBancos.Txt_key.Text)
    On Error GoTo 0
    pu_codcia = LK_CODCIA
    LEER_CCM_LLAVE
    If ccm_llave.EOF Then
            MENSAJE_BAN "Registro ,   NO EXISTE ... "
            Azul frmBancos.Txt_key, frmBancos.Txt_key
            GoTo fin
    End If
    LLENA_BAN 1
    frmBancos.Txt_key.Locked = True
    cmdCancelar.Enabled = True
    CmdModificar.SetFocus
    If Trim(TXTBAN(4).Text) <> "" Then
     SQ_OPER = 1
     PUB_CUENTA = Trim(TXTBAN(4).Text)
     LEER_COM_LLAVE
     If com_llave.EOF Then
       MsgBox "La Cuenta Contable: " & PUB_CUENTA & " Relacionada a este Banco a sido Eliminada en el Plan Contable, Favor de Relacionar nuevamente", 48, Pub_Titulo
       TXTBAN(4).Text = ""
       CmdModificar_Click ' activa la modificaion
       CmdModificar_Click ' graba la modificaion
       cmdcontab.Caption = "Re&lacionar"
     End If
    End If
    Screen.MousePointer = 0
Else
   If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   VALOR = UCase(ListView1.ListItems.Item(loc_key).Text)
   If Trim(UCase(Txt_key.Text)) = Left(VALOR, Len(Trim(Txt_key.Text))) Then
   Else
      Exit Sub
   End If
   ListView1.Visible = False
   cmdCancelar.Enabled = True
   LLENA_BAN 0
   If Trim(TXTBAN(4).Text) <> "" Then
     SQ_OPER = 1
     PUB_CUENTA = Trim(TXTBAN(4).Text)
     LEER_COM_LLAVE
     If com_llave.EOF Then
        MsgBox "La Cuenta Contable: " & PUB_CUENTA & " Relacionada a este Banco a sido Eliminada en el Plan Contable, Favor de Relacionar nuevamente", 48, Pub_Titulo
        TXTBAN(4).Text = ""
        CmdModificar_Click ' activa la modificaion
        CmdModificar_Click ' graba la modificaion
        cmdcontab.Caption = "Re&lacionar"
     End If
   End If

   frmBancos.Txt_key.Locked = True
   cmdCancelar.Enabled = True
   frmBancos.CmdModificar.SetFocus
End If
dale:
ListView1.Visible = False
fin:
mucho:

Exit Sub
SALCODI:
MsgBox Err.Description & " Intente Nuevamente ", 48, Pub_Titulo
Unload frmBancos
End Sub

Private Sub txt_key_KeyUp(KeyCode As Integer, Shift As Integer)
Dim NADA
Dim var
If Len(Txt_key.Text) = 0 Or IsNumeric(Txt_key.Text) = True Then
   ListView1.Visible = False
   Exit Sub
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(Txt_key.Text) = 1 Then
    If Txt_key.Text = "" Then Txt_key.Text = " "
    var = Asc(Txt_key.Text)
    var = var + 1
    NADA = var
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    numarchi = 9
    If LK_EMP = "3AA" Then
      archi = "SELECT * FROM CCMAEST WHERE  CCM_CODCIA = '" & par_llave!PAR_CIACCM & "' AND CCM_NOMBRE BETWEEN '" & Txt_key.Text & "' AND  '" & var & "' ORDER BY CCM_NOMBRE"
    Else
      archi = "SELECT * FROM CCMAEST WHERE  CCM_CODCIA = '" & LK_CODCIA & "' AND CCM_NOMBRE BETWEEN '" & Txt_key.Text & "' AND  '" & var & "' ORDER BY CCM_NOMBRE"
    End If
    
    PROC_LISVIEW ListView1
    loc_key = 1
    If NADA = 33 Or NADA = 91 Then
      If ListView1.Visible = False Then
        loc_key = 0
        MsgBox "No existe Datos ...", 48, Pub_Titulo
        Txt_key.Text = ""
      End If
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(Txt_key.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView1.ListItems.count Then
      ListView1.ListItems.Item(ListView1.ListItems.count).EnsureVisible
   Else
     ListView1.ListItems.Item(loc_key + 8).EnsureVisible
   End If
  End If
  Exit Sub
End If
End Sub


Private Sub Txtban_KeyPress(Index As Integer, KeyAscii As Integer)
Static ca As Integer
If KeyAscii = 27 And Index = 0 Then
   GoTo fin
End If
If KeyAscii = 13 And Index = 2 Then
     Azul ccm_nomcorto, ccm_nomcorto
     GoTo fin
End If
If KeyAscii = 13 And Index = 6 Then
    TXTBAN(7).SetFocus
     GoTo fin
End If


If KeyAscii = 13 And Index = 7 Then
    TXTBAN(4).SetFocus
     GoTo fin
End If

If KeyAscii = 13 And Index = 4 Then
    If CmdAgregar.Enabled Then
        CmdAgregar.SetFocus
    Else
        CmdModificar.SetFocus
    End If
    GoTo fin
End If

If KeyAscii = 13 And Index = 0 Then
   If Left(CmdAgregar.Caption, 2) = "&G" Or Left(CmdModificar.Caption, 2) = "&G" Then
     GoTo fin
   End If
    If Left(frmBancos.CmdCerrar.Caption, 3) = "&Ca" Then
        Exit Sub
    End If
    Dim VALOR As String
    Dim tf As Integer
    If frmBancos.Txt_key.Text = "" Then
        lista.Visible = True
        LEIDO.SetFocus
    Else
        SQ_OPER = 1
        PUB_CODBAN = Val(frmBancos.Txt_key.Text)
        pu_codcia = LK_CODCIA
        LEER_CCM_LLAVE
        If ccm_llave.EOF Then
                MENSAJE_BAN "Registro ,   NO EXISTE ... "
                Azul frmBancos.Txt_key, frmBancos.Txt_key
                GoTo fin
        End If
        LLENA_BAN 1
        frmBancos.Txt_key.Locked = True
        cmdCancelar.Enabled = True
        CmdModificar.SetFocus
   End If
End If

fin:
    
End Sub
Public Sub GRABA_CONTAB(wcia As String)
Exit Sub
If Trim(frmBancos.TXTBAN(4).Text) <> "" Then
    If Left(cmdcontab.Caption, 2) <> "&Q" Then
    Else
      'SQ_OPER = 1
      'PUB_CUENTA = Trim(TXTBAN(4).text)
      'LEER_COM_LLAVE
      'If Not com_llave.EOF Then
       '' If Trim(com_llave!com_descripcion) <> Trim(TXTBAN(2).text) Then
       '    pub_mensaje = "Desea tambien la Descripción de la Cta. Contable. ?"
       '    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
       '    If Pub_Respuesta = vbYes Then
       '         com_llave.Edit
       '         com_llave!com_descripcion = Trim(TXTBAN(2).text)
       '         com_llave.Update
       '    End If
       ' End If
      'End If
      Exit Sub
    End If
    com_llave.AddNew
    com_llave!COM_CODCIA = wcia
    com_llave!com_cuenta = LOC_CTA_CLI
    com_llave!com_descripcion = LOC_DES_CLI
    com_llave!COM_NIVEL = LOC_NIVEL
    com_llave!com_cuenta_sup = LOC_CTA_SUP
    com_llave!COM_FLAG_AFECTACION = LOC_FLAG_AFEC
    com_llave!com_ESTADO = LOC_ESTADO
    com_llave!COM_TIPO_CTA = LOC_TIPO_CTA
    com_llave!com_signo_d = LOC_SIGNO_D
    com_llave!com_signo_h = LOC_SIGNO_H
    com_llave!COM_act_pas = LOC_ACT_PAS
    com_llave!com_signo_h = LOC_SIGNO_H
    com_llave!COM_act_pas = LOC_ACT_PAS
    com_llave!COM_DEB_MES = 0
    com_llave!COM_HAB_MES = 0
    com_llave!COM_HAB_ANO = 0
    com_llave!COM_DEB_ANO = 0
    com_llave!com_cuenta_AUTOM_D = " "
    com_llave!com_cuenta_AUTO_H = " "
    com_llave!COM_CUENTA_AUTOM_D2 = " "
    com_llave!COM_CUENTA_AUTOM_D3 = " "
    com_llave!COM_CUENTA_AUTOM_D4 = " "
    com_llave!COM_CUENTA_AUTOM_D5 = " "
    com_llave!COM_POR_AUTOM_D = 0
    com_llave!COM_POR_AUTOM_D2 = 0
    com_llave!COM_POR_AUTOM_D3 = 0
    com_llave!COM_POR_AUTOM_D4 = 0
    com_llave!COM_POR_AUTOM_D5 = 0
    com_llave!COM_CENTRO_COSTOS = " "
 com_llave.Update
 cmdcontab.Caption = "&Quitar"
End If
End Sub

  

Public Sub crear_caja_chica()
ccm_llave.AddNew
ccm_llave!CCM_CODBAN = 1
ccm_llave!CCM_CODCIA = LK_CODCIA
ccm_llave!CCM_NOMBRE = "Caja Chica en"
ccm_llave!CCM_SALDO = 0
ccm_llave!CCM_CUENTA_CONTAB = " "
ccm_llave!CCM_MONEDA = "S"
ccm_llave!CCM_SIMBOLO = "S/."
ccm_llave!CCM_ULTIMA_OPERAC = 0
ccm_llave!CCM_SAL_ANTERIOR = 0
ccm_llave.Update


End Sub
