VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOpcion 
   Caption         =   "Opciones"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   Icon            =   "frmOpcion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ce&rrar"
      Height          =   555
      Left            =   4560
      Picture         =   "frmOpcion.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1455
   End
   Begin TabDlg.SSTab tab 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   882
      ForeColor       =   8388608
      TabCaption(0)   =   "Almacenes"
      TabPicture(0)   =   "frmOpcion.frx":058C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdmes"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdano"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CHEBLO"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Cuentas Contables Constantes"
      TabPicture(1)   =   "frmOpcion.frx":05A8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text2"
      Tab(1).Control(1)=   "Text1"
      Tab(1).Control(2)=   "ctacon(1)"
      Tab(1).Control(3)=   "ctacon(0)"
      Tab(1).ControlCount=   4
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -73080
         TabIndex        =   12
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -73080
         TabIndex        =   11
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox CHEBLO 
         Caption         =   "Bloqueo"
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
         Height          =   255
         Left            =   4440
         TabIndex        =   5
         Top             =   1005
         Width           =   1335
      End
      Begin VB.ComboBox cmdano 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox cmdmes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label ctacon 
         Caption         =   "Cuenta de Ganacia por Tipo Cambio. :"
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   -74880
         TabIndex        =   10
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label ctacon 
         Caption         =   "Cuenta de Perdida por Tipo Cambio. :"
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   -74880
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Año"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Mes"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Bloqueo por Meses (Ninguna modificación que afecten el costo)"
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   855
         Width           =   1575
      End
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
      Left            =   -105
      TabIndex        =   8
      Top             =   4320
      Width           =   6375
   End
End
Attribute VB_Name = "frmOpcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WFLAG As String * 1
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset
Dim PS_REP02 As rdoQuery
Dim llave_rep02 As rdoResultset

Private Sub cmdano_Click()
Dim WNUMERO As String
WNUMERO = Trim(Right(cmdmes.Text, 8)) & Trim(cmdano.Text)
SQ_OPER = 1
PUB_TIPREG = 60
PUB_NUMTAB = val(WNUMERO)
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
WFLAG = "A"
CHEBLO.Value = 0
If Not tab_llave.EOF Then
  CHEBLO.Value = tab_llave!TAB_CODART
End If
WFLAG = ""
End Sub

Private Sub cmdmes_Click()
Dim WNUMERO As String
WNUMERO = Trim(Right(cmdmes.Text, 8)) & Trim(cmdano.Text)
SQ_OPER = 1
PUB_TIPREG = 60
PUB_NUMTAB = val(WNUMERO)
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
WFLAG = "A"
If Not tab_llave.EOF Then
  CHEBLO.Value = tab_llave!TAB_CODART
Else
  CHEBLO.Value = 0
End If
WFLAG = ""

End Sub

Private Sub Command1_Click()
Unload frmOpcion
End Sub

Private Sub CHEBLO_Click()
WNUMERO = Trim(Right(cmdmes.Text, 8)) & Trim(cmdano.Text)
If WFLAG = "A" Then Exit Sub
If tab_llave.EOF Then
 tab_llave.AddNew
Else
 tab_llave.Edit
End If
tab_llave!TAB_CODCIA = LK_CODCIA
tab_llave!TAB_TIPREG = 60
tab_llave!TAB_NUMTAB = val(WNUMERO)
tab_llave!TAB_CODART = CHEBLO.Value
tab_llave.Update
'tab_llave.CancelUpdate

End Sub

Private Sub Form_Activate()
Dim wmes
Dim wano
wmes = Format(LK_FECHA_DIA, "mm")
wano = Format(LK_FECHA_DIA, "yyyy")
For fila = 0 To cmdmes.ListCount - 1
 cmdmes.ListIndex = fila
 If val(Right(cmdmes.Text, 8)) = val(wmes) Then
   Exit For
 End If
Next fila
For fila = 0 To cmdano.ListCount - 1
 cmdano.ListIndex = fila
 If val(cmdano.Text) = val(wano) Then
   Exit For
 End If
Next fila

End Sub

Private Sub Form_Load()
WFLAG = ""
cmdmes.AddItem "Enero " & String(60, " ") & "01"
cmdmes.AddItem "Febrero " & String(60, " ") & "02"
cmdmes.AddItem "Marzo " & String(60, " ") & "03"
cmdmes.AddItem "Abril " & String(60, " ") & "04"
cmdmes.AddItem "Mayo " & String(60, " ") & "05"
cmdmes.AddItem "Junio " & String(60, " ") & "06"
cmdmes.AddItem "Julio " & String(60, " ") & "07"
cmdmes.AddItem "Agosto " & String(60, " ") & "08"
cmdmes.AddItem "Septiembre " & String(60, " ") & "09"
cmdmes.AddItem "Octubre " & String(60, " ") & "10"
cmdmes.AddItem "Noviembre " & String(60, " ") & "11"
cmdmes.AddItem "Diciembre " & String(60, " ") & "12"
cmdano.AddItem "2000"
cmdano.AddItem "2001"
cmdano.AddItem "2002"
cmdano.AddItem "2003"
cmdano.AddItem "2004"
cmdano.AddItem "2005"
cmdano.AddItem "2006"

End Sub

