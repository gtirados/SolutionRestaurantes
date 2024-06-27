VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTC 
   Caption         =   "Tipo de Cambios"
   ClientHeight    =   6036
   ClientLeft      =   2568
   ClientTop       =   1920
   ClientWidth     =   4380
   Icon            =   "frmtc.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6036
   ScaleWidth      =   4380
   Begin VB.CommandButton Command2 
      Caption         =   "Ce&rrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   4215
      Begin RichTextLib.RichTextBox TEXTOVAR 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1715
         _ExtentY        =   656
         _Version        =   327680
         BackColor       =   16776960
         BorderStyle     =   0
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"frmtc.frx":0442
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         _ExtentX        =   7006
         _ExtentY        =   7006
         _Version        =   327680
         Rows            =   3
         FixedRows       =   2
      End
      Begin MSMask.MaskEdBox txtfecha 
         Height          =   285
         Left            =   2520
         TabIndex        =   0
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2350
         _ExtentY        =   508
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtfechaI 
         Height          =   285
         Left            =   480
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2350
         _ExtentY        =   508
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
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
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   9
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "De"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label3 
         Caption         =   "Ingrese Tipo de Cambio :"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   15
      X1              =   0
      X2              =   4440
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lblcierre 
      AutoSize        =   -1  'True
      Caption         =   "Act. de Tipo de Cambio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   720
      TabIndex        =   6
      Top             =   0
      Width           =   2805
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
CenterMe frmTC
txtfecha.Text = Format(LK_FECHA_COP2, "dd/mm/yyyy")
txtfecha.Mask = "##/##/####"
txtfecha.Visible = True
txtfecha.TabIndex = 1

txtfechaI.Text = Format(LK_FECHA_COP1, "dd/mm/yyyy")
txtfechaI.Mask = "##/##/####"
txtfechaI.Visible = True
txtfechaI.TabIndex = 0
Muestra_tc txtfechaI.Text, txtfecha.Text
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

Public Sub Muestra_tc(wfecha_ini As Date, wfecha_Fin As Date)
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
PUB_CAL_FIN = wfecha_Fin 'wdiaF & "/" & wmesM & "/" & PUB_CAL_ANO
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
If Right(txtfecha.Text, 2) = "__" Then
   txtfecha.Text = Format(Left(txtfecha.Text, 8), "dd/mm/yyyy")
End If

If Not IsDate(txtfecha.Text) Then
  MsgBox "Fecha no procede. ", 48, Pub_Titulo
  Azul2 txtfecha, txtfecha
  Exit Sub
End If
'If txtfecha.Text > LK_FECHA_DIA Then
'  MsgBox "Fecha no Puede ser mayor a la del dia.", 48, Pub_Titulo
'  Exit Sub
'End If
temfecha = txtfecha.Text

Muestra_tc txtfechaI.Text, txtfecha.Text
End If
End Sub

Public Function JALAR(wfecha_ini As Date, wfecha_Fin As Date) As Currency
PUB_CAL_INI = wfecha_ini
PUB_CAL_FIN = wfecha_Fin
pu_codcia = LK_CODCIA
PUB_CODCIA = LK_CODCIA
SQ_OPER = 1
LEER_CAL_LLAVE
JALAR = cal_llave!cal_tipo_cambio
End Function

Private Sub txtfechaI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Right(txtfechaI.Text, 2) = "__" Then
     txtfechaI.Text = Format(Left(txtfechaI.Text, 8), "dd/mm/yyyy")
  End If
  If Not IsDate(txtfechaI) Then
    Azul2 txtfechaI, txtfechaI
    Exit Sub
  End If
  Azul2 txtfecha, txtfecha
End If
End Sub
