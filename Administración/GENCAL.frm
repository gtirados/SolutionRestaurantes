VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form GENCAL 
   BackColor       =   &H00F5F1EC&
   Caption         =   "Calendario del Sistema"
   ClientHeight    =   3600
   ClientLeft      =   1485
   ClientTop       =   2220
   ClientWidth     =   8595
   ClipControls    =   0   'False
   ForeColor       =   &H00F5F1EC&
   Icon            =   "GENCAL.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3600
   ScaleWidth      =   8595
   Begin VB.CommandButton Command3 
      Caption         =   "Generar Cal. T.C."
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
      Left            =   2520
      TabIndex        =   22
      Top             =   2880
      Width           =   2055
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
      Left            =   6240
      TabIndex        =   19
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00F5F1EC&
      Height          =   1935
      Left            =   4800
      TabIndex        =   13
      Top             =   600
      Width           =   3615
      Begin VB.TextBox TxtAno 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   735
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   300
         Left            =   2520
         TabIndex        =   17
         Top             =   480
         Width           =   230
      End
      Begin VB.CommandButton CmdVerCalen 
         Caption         =   "&Mostrar Calendario"
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
         Left            =   720
         TabIndex        =   15
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione Año a Editar :"
         ForeColor       =   &H009C3000&
         Height          =   435
         Left            =   225
         TabIndex        =   16
         Top             =   360
         Width           =   1305
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F5F1EC&
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4575
      Begin VB.CommandButton Command1 
         Caption         =   "&Procesar Calendario "
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
         Left            =   135
         TabIndex        =   9
         Top             =   2265
         Width           =   2055
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00F5F1EC&
         Caption         =   "Fechas a Procesar  : "
         Height          =   1215
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   3975
         Begin VB.CheckBox ChkMarcar 
            BackColor       =   &H00F5F1EC&
            Caption         =   "Marcar los Feriados y Domingos"
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
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Value           =   1  'Checked
            Width           =   3135
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   300
            Left            =   1440
            TabIndex        =   6
            Top             =   360
            Width           =   230
         End
         Begin VB.TextBox TxtFecIni 
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
            Left            =   720
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   5
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox TxtFecFin 
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
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   360
            Width           =   735
         End
         Begin VB.VScrollBar VScroll2 
            Height          =   300
            Left            =   3240
            TabIndex        =   3
            Top             =   360
            Width           =   230
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "De"
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
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "a"
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
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   7
            Top             =   360
            Width           =   255
         End
      End
      Begin ComctlLib.ProgressBar ProgBar 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label POR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Generando..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   1560
         TabIndex        =   21
         Top             =   1800
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label POR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3960
         TabIndex        =   12
         Top             =   1920
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label POR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Editar Calendario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009C3000&
      Height          =   330
      Left            =   5400
      TabIndex        =   14
      Top             =   120
      Width           =   2430
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Generación de Calendario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009C3000&
      Height          =   330
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3585
   End
End
Attribute VB_Name = "GENCAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdVerCalen_Click()
Screen.MousePointer = 11
If Trim(TxtAno.text) = "" Then
  Screen.MousePointer = 0
  Exit Sub
End If
PUB_CAL_ANO = Trim(TxtAno.text)
GENCAL.Hide
Load FrmGrafCal
FrmGrafCal.Show 1
End Sub

Private Sub Command1_Click()
Dim CAL As rdoResultset
Dim fecha2 As Date
Dim CONTADOR As Integer
Dim WS_MESDIA As String
Dim WS_DIA As String
Dim WS_MES As String
Dim WS_ANOS As Integer
Dim WS_TOTDIAS As Integer
Dim wsfeci As Integer
Dim wsfecf As Integer
pub_mensaje = "Agregar Calendario  ¿Desea adicionar Calendario... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If
Screen.MousePointer = 11
Command1.Enabled = False
pub_cadena = "SELECT * FROM CALENDARIO ORDER BY CAL_FECHA"
Set CAL = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)

 '  pub_cadena = "DELETE CALENDARIO WHERE CAL_CODCIA = '" & LK_CODCIA & "'"
 '  CN.Execute pub_cadena, rdExecDirect


CONTADOR = 0
WS_ANOS = Val(TxtFecFin) - Val(TxtFecIni) + 1
WS_TOTDIAS = WS_ANOS * 365
ProgBar.Min = 1
ProgBar.Max = WS_TOTDIAS
DoEvents
ProgBar.Visible = True
DoEvents
POR(0).Visible = True
DoEvents
POR(1).Visible = True
DoEvents
POR(2).Visible = True
DoEvents

fecha2 = "01/1/" & TxtFecIni.text
Do Until CONTADOR = WS_TOTDIAS
   CAL.AddNew
   CAL!CAL_CODCIA = LK_CODCIA
   CAL!CAL_LABORABLE = "S"
   CAL!CAL_INDICE = "0"
   CAL!CAL_FECHA = fecha2
   WS_DIA = DatePart("d", fecha2)
   WS_MES = DatePart("m", fecha2)
   WS_MESDIA = WS_MES & WS_DIA
   WS_DIA = DatePart("w", fecha2)
   If WS_DIA = "1" And ChkMarcar.Value = 1 Then
      CAL!CAL_LABORABLE = "N"
   End If
   CAL!CAL_TIPO_CAMBIO = 0
   CAL!CAL_TC_MERCA = 0
   CAL.Update
   fecha2 = DateAdd("d", 1, fecha2)
   CONTADOR = CONTADOR + 1
   'DoEvents
   ProgBar.Value = CONTADOR
Loop
CAL.MoveFirst
CAL.Edit
CAL!CAL_INDICE = "1"
CAL.Update
'***** FERIADOS
If ChkMarcar.Value <> 1 Then
  GoTo sigue
End If
CONTADOR = 0
CAL.MoveFirst
wsfeci = DatePart("yyyy", CAL!CAL_FECHA)
CAL.MoveLast
wsfecf = DatePart("yyyy", CAL!CAL_FECHA)

pub_cadena = "SELECT * FROM calfer ORDER BY FER_FECHA "
Set PSFER_LLAVE = CN.CreateQuery("", pub_cadena)
Set fer_llave = PSFER_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

fer_llave.Requery
fer_llave.MoveFirst
SQ_OPER = 1
Do Until fer_llave.EOF
  For CONTADOR = wsfeci To wsfecf
      WS_DIA = DatePart("d", fer_llave!FER_FECHA) & "/" & DatePart("m", fer_llave!FER_FECHA) & "/" & CONTADOR
      PUB_FECHA = WS_DIA
      PUB_CAL_INI = PUB_FECHA
      PUB_CAL_FIN = PUB_FECHA
      PUB_CODCIA = LK_CODCIA
      LEER_CAL_LLAVE
      If Not cal_llave.EOF Then
       cal_llave.Edit
       cal_llave!CAL_LABORABLE = "N"
       cal_llave.Update
      End If
  Next CONTADOR
  fer_llave.MoveNext
Loop

sigue:
PUB_CAL_INI = 0
PUB_CAL_FIN = #1/1/2050#
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_CAL_LLAVE 1
If cal_llave.EOF Then
    MsgBox "NO Tiene Calendario el Sistema ...Generarlo..!!!!!", 46, Pub_Titulo
    GENCAL.CmdVerCalen.Enabled = False
Else
   cal_llave.MoveFirst
   wsfeci = DatePart("yyyy", cal_llave!CAL_FECHA)
   cal_llave.MoveLast
   wsfecf = DatePart("yyyy", cal_llave!CAL_FECHA)
   VScroll3.LargeChange = 1
   VScroll3.SmallChange = 1
   VScroll3.Min = wsfecf
   VScroll3.Max = wsfeci
   VScroll3.Value = wsfeci
   TxtAno.text = wsfeci
End If
ProgBar.Visible = False
POR(0).Visible = False
POR(1).Visible = False
POR(2).Visible = False
Screen.MousePointer = 0
Command1.Enabled = True
End Sub


Private Sub Command2_Click()
Unload GENCAL
End Sub

Private Sub Command3_Click()
Dim CAL As rdoResultset
Dim fecha2 As Date
Dim CONTADOR As Integer
Dim WS_MESDIA As String
Dim WS_DIA As String
Dim WS_MES As String
Dim WS_ANOS As Integer
Dim WS_TOTDIAS As Integer
Dim wsfeci As Integer
Dim wsfecf As Integer
On Error GoTo sigue2
pub_mensaje = " ¿Desea Generar el Nuevo Calendario... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If
Screen.MousePointer = 11
Command1.Enabled = False
pub_cadena = "SELECT * FROM CALENDARIO ORDER BY CAL_FECHA"
Set CAL = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
CONTADOR = 0
WS_ANOS = Val(TxtFecFin) - Val(TxtFecIni) + 1
WS_TOTDIAS = WS_ANOS * 365
ProgBar.Min = 1
ProgBar.Max = WS_TOTDIAS
DoEvents
ProgBar.Visible = True
DoEvents
POR(0).Visible = True
DoEvents
POR(1).Visible = True
DoEvents
POR(2).Visible = True
DoEvents

fecha2 = "01/1/" & TxtFecIni.text
Do Until CONTADOR = WS_TOTDIAS
   CAL.AddNew
   CAL!CAL_CODCIA = "00"
   CAL!CAL_LABORABLE = "S"
   CAL!CAL_INDICE = "0"
   CAL!CAL_FECHA = fecha2
   WS_DIA = DatePart("d", fecha2)
   WS_MES = DatePart("m", fecha2)
   WS_MESDIA = WS_MES & WS_DIA
   WS_DIA = DatePart("w", fecha2)
   If WS_DIA = "1" And ChkMarcar.Value = 1 Then
      CAL!CAL_LABORABLE = "N"
   End If
   CAL!CAL_TIPO_CAMBIO = 0
   CAL!CAL_TC_MERCA = 0
   CAL.Update
   fecha2 = DateAdd("d", 1, fecha2)
   CONTADOR = CONTADOR + 1
   'DoEvents
   ProgBar.Value = CONTADOR
Loop
CAL.MoveFirst
CAL.Edit
CAL!CAL_INDICE = "1"
CAL.Update
'***** FERIADOS
If ChkMarcar.Value <> 1 Then
  GoTo sigue
End If
CONTADOR = 0
CAL.MoveFirst
wsfeci = DatePart("yyyy", CAL!CAL_FECHA)
CAL.MoveLast
wsfecf = DatePart("yyyy", CAL!CAL_FECHA)

sigue:
PUB_CAL_INI = 0
PUB_CAL_FIN = #1/1/2050#
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_CAL_LLAVE
If cal_llave.EOF Then
    MsgBox "NO Tiene Calendario el Sistema ...Generarlo..!!!!!", 46, Pub_Titulo
    GENCAL.CmdVerCalen.Enabled = False
Else
   cal_llave.MoveFirst
   wsfeci = DatePart("yyyy", cal_llave!CAL_FECHA)
   cal_llave.MoveLast
   wsfecf = DatePart("yyyy", cal_llave!CAL_FECHA)
   VScroll3.LargeChange = 1
   VScroll3.SmallChange = 1
   VScroll3.Min = wsfecf
   VScroll3.Max = wsfeci
   VScroll3.Value = wsfeci
   TxtAno.text = wsfeci
End If
ProgBar.Visible = False
POR(0).Visible = False
POR(1).Visible = False
POR(2).Visible = False
Screen.MousePointer = 0
Command1.Enabled = True

sigue2:

Resume Next
End Sub

Private Sub Form_Load()
    Dim wsfeci As Integer
    Dim wsfecf As Integer
    
PUB_CAL_INI = #1/1/1900#
PUB_CAL_FIN = #12/31/2050#
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_CAL_LLAVE 1
    cal_llave.Requery
    If cal_llave.EOF Then
        MsgBox "Esta Empresa NO Tiene su Calendario  ...Generarlo..!!!!!", 48, Pub_Titulo
        'GENCAL.CmdVerCalen.Enabled = False
    Else
        cal_llave.MoveFirst
        wsfeci = DatePart("yyyy", cal_llave!CAL_FECHA)
        cal_llave.MoveLast
        wsfecf = DatePart("yyyy", cal_llave!CAL_FECHA)
        VScroll3.LargeChange = 1
        VScroll3.SmallChange = 1
        VScroll3.Min = wsfecf
        VScroll3.Max = wsfeci
        VScroll3.Value = wsfeci
        TxtAno.text = wsfeci
    End If
    
    VScroll1.Min = 2100
    VScroll1.Max = 1900
    VScroll1.LargeChange = 1
    VScroll1.SmallChange = 1
    VScroll1 = Year(Date)
    VScroll2.LargeChange = 1
    VScroll2.SmallChange = 1
    CenterMe GENCAL
Command3.Visible = False
If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Then
  Command3.Visible = True
End If
End Sub


Private Sub VScroll1_Change()
 TxtFecIni.text = VScroll1.Value
 VScroll2.Min = VScroll1.Value + 11
 VScroll2.Max = VScroll1.Value
 VScroll2.Value = VScroll1.Value
End Sub

Private Sub VScroll2_Change()
TxtFecFin.text = VScroll2.Value
End Sub

Private Sub VScroll3_Change()
If Trim(TxtAno.text) = "" Then Exit Sub
TxtAno.text = VScroll3.Value
End Sub
