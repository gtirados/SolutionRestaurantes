VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTC 
   BackColor       =   &H00FAEFDA&
   Caption         =   "Tipo de Cambios"
   ClientHeight    =   6030
   ClientLeft      =   2565
   ClientTop       =   1920
   ClientWidth     =   4380
   Icon            =   "frmtc.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   4380
   Begin VB.CommandButton Command2 
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
      Height          =   450
      Left            =   3075
      TabIndex        =   2
      Top             =   5520
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4905
      Left            =   75
      TabIndex        =   3
      Top             =   570
      Width           =   4215
      Begin RichTextLib.RichTextBox TEXTOVAR 
         Height          =   300
         Left            =   1680
         TabIndex        =   4
         Top             =   1365
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393217
         BackColor       =   16445402
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"frmtc.frx":0442
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid gridigv 
         Height          =   3975
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "[Enter] = para Editar"
         Top             =   840
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   7011
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
      End
      Begin MSMask.MaskEdBox txtfecha 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pulsar {Enter}"
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
         Left            =   2520
         TabIndex        =   8
         Top             =   360
         Width           =   1230
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese Tipo de Cambio :"
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
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "A partir de:"
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   60
      Top             =   5550
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   15
      X1              =   0
      X2              =   4440
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Label lblcierre 
      AutoSize        =   -1  'True
      BackColor       =   &H00FAEFDA&
      Caption         =   "Act. de Tipo de Cambio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   7
      Top             =   0
      Width           =   2835
   End
End
Attribute VB_Name = "frmTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim temporal
Dim temfecha


Private Sub Command2_Click()
Unload frmTC
End Sub

Private Sub Form_Activate()
If txtfecha.Visible Then
 txtfecha_KeyPress 13
 Azul2 txtfecha, txtfecha
End If
End Sub

Private Sub Form_Load()
'CenterMe Costos

txtfecha.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
txtfecha.Mask = "##/##/####"
txtfecha.Visible = True
txtfecha.TabIndex = 0
Muestra_tc txtfecha.Text
temfecha = LK_FECHA_DIA
End Sub
Private Sub gridiGV_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, WC
Static CONS
If KeyAscii <> 13 Then Exit Sub

'If Trim(gridigv.TextMatrix(gridigv.Row, 9)) <> "8" Then
'  If Trim(gridigv.TextMatrix(gridigv.Row, 0)) = "" Then Exit Sub
'  If Trim(gridigv.TextMatrix(gridigv.Row, 1)) <> "" And gridigv.Col = 2 Or gridigv.Col = 3 Then GoTo leer
'  If Trim(gridigv.TextMatrix(gridigv.Row, 8)) <> "0" Then Exit Sub
'End If


'If gridigv.Col = 1 And WMODO = "I" Then
'   a = Val(gridigv.TextMatrix(gridigv.Row - 1, 0))
'   a = a + 1
'  gridigv.TextMatrix(gridigv.Row, 0) = a
'End If
'If WMODO = "I" Or WMODO = "C" Then
    TEXTOVAR.Left = gridigv.Left + gridigv.CellLeft
    TEXTOVAR.Width = gridigv.CellWidth
    TEXTOVAR.Height = gridigv.CellHeight
    TEXTOVAR.Top = gridigv.Top + gridigv.CellTop
    TEXTOVAR.Text = gridigv.TextMatrix(gridigv.Row, gridigv.Col)
    TEXTOVAR.Visible = True
    Azul3 TEXTOVAR, TEXTOVAR
    TEXTOVAR.SetFocus
'End If
End Sub

Private Sub gridiGV_KeyUp(KeyCode As Integer, Shift As Integer)
Dim WC
Dim a, WF As Integer
Dim tf, t, tc
Dim SALE As Boolean
Dim Wsec

'If WMODO = "C" Then Exit Sub

'If cop_llave!COP_FLAG_MAYORIZACION = "M" Then
 'MsgBox "Ojo estaba Mayorizado..."
'End If


If Left(gridigv.TextMatrix(gridigv.Row, 0), 2) <> "MA" Then Exit Sub
 If KeyCode = 32 Then
  'If WMODO <> "C" Then Exit Sub
  tc = gridigv.Col
  For fila = 1 To gridigv.Cols - 1
      gridigv.Col = fila
      If gridigv.CellBackColor = QBColor(12) Then
         gridigv.CellBackColor = QBColor(15)
         gridigv.TextMatrix(gridigv.Row, 9) = "9"
      Else
         gridigv.CellBackColor = QBColor(12)
         gridigv.TextMatrix(gridigv.Row, 9) = "-1"
      End If
  Next fila
  gridigv.Col = tc
  gridigv.SetFocus
  Exit Sub
End If
If KeyCode = 45 Then
    Wsec = Wsec + 1
    If Trim(gridigv.TextMatrix(gridigv.Row + 1, 11)) = "8" Then
         Exit Sub
    Else
      If Trim(gridigv.TextMatrix(gridigv.Row + 1, 0)) = "T" Then Exit Sub
    End If
    If Val(gridigv.TextMatrix(gridigv.Row, 4)) = 0 And Val(gridigv.TextMatrix(gridigv.Row, 5)) = 0 Then Exit Sub
    gridigv.AddItem "", gridigv.Row + 1
    gridigv.TextMatrix(gridigv.Row + 1, 0) = "MAN. " & Format(gridigv.TextMatrix(gridigv.Row, 10), "dd/mm/yyyy")
    gridigv.TextMatrix(gridigv.Row + 1, 6) = Wsec
    gridigv.TextMatrix(gridigv.Row + 1, 8) = gridigv.TextMatrix(gridigv.Row, 8)
    gridigv.TextMatrix(gridigv.Row + 1, 3) = gridigv.TextMatrix(gridigv.Row, 3)
    gridigv.TextMatrix(gridigv.Row + 1, 7) = gridigv.TextMatrix(gridigv.Row, 7)
    gridigv.TextMatrix(gridigv.Row + 1, 10) = gridigv.TextMatrix(gridigv.Row, 10)
    gridigv.TextMatrix(gridigv.Row + 1, 11) = "8"
    gridigv.Row = gridigv.Row + 1
    gridigv.Col = 1
    gridigv.SetFocus
End If
Exit Sub
If KeyCode = 46 Then
If gridigv.Rows <= 3 Then
Else
   pub_mensaje = MsgBox("Desea Quitar el Item de la Cuenta : " & Trim(gridigv.TextMatrix(gridigv.Row, 1)), vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
   If pub_mensaje = vbNo Then
     gridigv.SetFocus
     Exit Sub
   Else
     gridigv.RowHeight(gridigv.Row) = 1
     gridigv.Row = gridigv.Row + 1
    
   'gridiGV.RemoveItem (gridiGV.Row)
   'gridiGV.Refresh
   gridigv.SetFocus
   End If
End If
End If
'gridiGV.SetFocus
Exit Sub



End Sub

Private Sub gridigv_Scroll()
TEXTOVAR.Visible = False
End Sub


Private Sub textovar_Change()
gridigv.Text = Format(TEXTOVAR.Text, "0.0000000")
End Sub

Private Sub TEXTOVAR_GotFocus()
 temporal = gridigv.TextMatrix(gridigv.Row, gridigv.Col)
End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  TEXTOVAR.Text = temporal
  TEXTOVAR.Visible = False
  gridigv.SetFocus
  Exit Sub
End If
If gridigv.Col = 1 Then Consistencias gridigv, TEXTOVAR, KeyAscii
If gridigv.Col = 4 Then Consistencias gridigv, TEXTOVAR, KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If
If gridigv.Col = 1 Or gridigv.Col = 4 Then
  If Val(TEXTOVAR.Text) > 99 Then
    Azul3 TEXTOVAR, TEXTOVAR
    Exit Sub
  End If
End If

PUB_CAL_INI = gridigv.TextMatrix(gridigv.Row, 2)
PUB_CAL_FIN = gridigv.TextMatrix(gridigv.Row, 2)
pu_codcia = LK_CODCIA
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_CAL_LLAVE
cal_llave.Edit
If gridigv.Col = 1 Then
   cal_llave!cal_tipo_cambio = Val(TEXTOVAR.Text)
End If
If gridigv.Col = 4 Then
   cal_llave!CAL_TC_MERCA = Val(TEXTOVAR.Text)
   If Format(LK_FECHA_DIA, "dd/mm/yyyy") = Format(gridigv.TextMatrix(gridigv.Row, 0), "dd/mm/yyyy") Then
      LK_TIPO_CAMBIO = Val(TEXTOVAR.Text)
      MDIForm1.StatusBar1.Panels(3).Text = "T.C.= S/. " + Format(LK_TIPO_CAMBIO, "0.0000")
   End If
End If
cal_llave.Update
If gridigv.Row >= gridigv.Rows - 1 Then
Else
  gridigv.Row = gridigv.Row + 1
End If
gridigv.SetFocus
TEXTOVAR.Visible = False

fin:

End Sub


Private Sub Timer1_Timer()
'lblcierre.Visible = Not lblcierre.Visible
End Sub

Public Sub Muestra_tc(wfecha_ini As Date)
Dim wdiaI, wdiaF As String
Dim wmesM As String
gridigv.Clear
gridigv.Cols = 5
gridigv.Rows = 2
gridigv.ColWidth(0) = 1000
gridigv.ColWidth(1) = 1200
gridigv.ColWidth(2) = 0
gridigv.ColWidth(3) = 0
gridigv.ColWidth(4) = 1200

gridigv.TextMatrix(0, 0) = "Fecha"
gridigv.TextMatrix(0, 1) = "T.Cambio"
gridigv.TextMatrix(1, 0) = "-"
gridigv.Row = 1
gridigv.Col = 1
gridigv.TextMatrix(1, 1) = "Venta"
gridigv.CellForeColor = QBColor(9)
gridigv.CellFontBold = True
gridigv.CellAlignment = 4
gridigv.TextMatrix(0, 4) = "T.Cambio"
gridigv.Col = 4
gridigv.TextMatrix(1, 4) = "Compra"
gridigv.CellForeColor = QBColor(4)
gridigv.CellFontBold = True
gridigv.CellAlignment = 4


PUB_CAL_INI = wfecha_ini ' wdiaI & "/" & wmesM & "/" & PUB_CAL_ANO
PUB_CAL_FIN = LK_FECHA_DIA 'wdiaF & "/" & wmesM & "/" & PUB_CAL_ANO
pu_codcia = LK_CODCIA
PUB_CODCIA = LK_CODCIA
SQ_OPER = 1
LEER_CAL_LLAVE
cal_llave.MoveFirst
fila = 1
gridigv.Rows = 2
Do Until cal_llave.EOF
  fila = fila + 1
  gridigv.Rows = gridigv.Rows + 1
  gridigv.RowHeight(gridigv.Rows - 1) = 285
  gridigv.TextMatrix(fila, 0) = Format(cal_llave!CAL_FECHA, "dd/mm/yyyy")
  gridigv.TextMatrix(fila, 1) = Format(cal_llave!cal_tipo_cambio, "0.0000000")
  gridigv.TextMatrix(fila, 2) = Format(cal_llave!CAL_FECHA, "dd/mm/yyyy")
  gridigv.TextMatrix(fila, 4) = Format(cal_llave!CAL_TC_MERCA, "0.0000000")
  cal_llave.MoveNext
Loop
gridigv.Visible = True
gridigv.Col = 1
gridigv.Row = 2
If gridigv.Visible Then gridigv.SetFocus

End Sub

Public Function CONSIS_TC(wfecha_ini As Date) As Boolean
Dim WRES As String
PUB_CAL_INI = wfecha_ini ' wdiaI & "/" & wmesM & "/" & PUB_CAL_ANO
PUB_CAL_FIN = LK_FECHA_DIA 'wdiaF & "/" & wmesM & "/" & PUB_CAL_ANO
pu_codcia = LK_CODCIA
PUB_CODCIA = LK_CODCIA
SQ_OPER = 1
LEER_CAL_LLAVE
WRES = "Las Siguientes Fechas no tienen tipo de cambio:" & Chr(13)
fila = 0
Do Until cal_llave.EOF
If Nulo_Valor0(cal_llave!cal_tipo_cambio) = 0 Then
  WRES = WRES + Format(cal_llave!CAL_FECHA, "dd/mm/yyyy") + " "
  fila = 1
End If
cal_llave.MoveNext
Loop
WRES = WRES + "." + Chr(13) + "Consulte su tabla de tipo de cambios"
If fila = 1 Then
  MsgBox WRES, 48, Pub_Titulo
  Muestra_tc wfecha_ini
  CONSIS_TC = False
Else
   CONSIS_TC = True
End If

End Function
Private Sub Consistencias(wsGrid As MSFlexGrid, wsTexto As RichTextBox, wsKeyAscii As Integer)
  Static valor
  Dim car As String
 ' NUMEROS CON DECIMALES
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If wsKeyAscii = 45 Then
      If wsTexto.Text <> "" Then
         Beep
         wsKeyAscii = 0
         Exit Sub
      End If
    End If
    If wsKeyAscii = 46 Then
      If InStr(1, wsTexto.Text, ".") <> 0 Then
        Beep
        wsKeyAscii = 0
        Exit Sub
      End If
    End If
    
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And car <> "." Then
          wsKeyAscii = 0
          Beep
          Exit Sub
        End If
    End If

End Sub

Private Sub txtfecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Not IsDate(txtfecha.Text) Then
  MsgBox "Fecha no procede. ", 48, Pub_Titulo
  Azul2 txtfecha, txtfecha
  Exit Sub
End If
If txtfecha.Text > LK_FECHA_DIA Then
  MsgBox "Fecha no Puede ser mayor a la del dia.", 48, Pub_Titulo
  Exit Sub
End If
temfecha = txtfecha.Text

Muestra_tc txtfecha.Text
End If
End Sub

Public Function JALAR(wfecha_ini As Date, wfecha_fin As Date) As Currency
PUB_CAL_INI = wfecha_ini
PUB_CAL_FIN = wfecha_fin
pu_codcia = LK_CODCIA
PUB_CODCIA = LK_CODCIA
SQ_OPER = 1
LEER_CAL_LLAVE
JALAR = cal_llave!cal_tipo_cambio
End Function
