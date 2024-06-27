VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCuotas 
   Caption         =   "Definición de Cuotas."
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   6615
      Left            =   30
      TabIndex        =   6
      Top             =   750
      Width           =   11655
      Begin RichTextLib.RichTextBox TEXTOVAR 
         Height          =   375
         Left            =   6240
         TabIndex        =   20
         Top             =   1890
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393217
         BackColor       =   12632064
         BorderStyle     =   0
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"frmCuotas.frx":0000
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
      Begin VB.ComboBox cmbdivi 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox Txt_key 
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
         Left            =   6855
         MaxLength       =   10
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox i_codart2 
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
         Left            =   960
         MaxLength       =   8
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdava 
         Caption         =   "Mostrar Avance"
         Height          =   435
         Left            =   2400
         TabIndex        =   9
         Top             =   5520
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Grabar"
         Height          =   435
         Left            =   600
         TabIndex        =   8
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton Cmdclose 
         Caption         =   "Ce&rrar"
         Height          =   435
         Left            =   9480
         TabIndex        =   10
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton cmdmostrar 
         Caption         =   "&Mostrar"
         Height          =   375
         Left            =   5565
         TabIndex        =   7
         Top             =   690
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid gridigv 
         Height          =   3795
         Left            =   135
         TabIndex        =   19
         ToolTipText     =   "[Enter] = para Editar"
         Top             =   1440
         Width           =   11370
         _ExtentX        =   20055
         _ExtentY        =   6694
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         BackColorBkg    =   16777215
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   -165
         X2              =   11655
         Y1              =   1215
         Y2              =   1215
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   11640
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "División :"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo :"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   6000
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo :"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Detalle de Cuotas :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   135
         TabIndex        =   12
         Top             =   1260
         Width           =   1350
      End
   End
   Begin VB.Frame fratipo 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   11475
      Begin VB.ComboBox Cmbtipos 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   4215
      End
      Begin MSMask.MaskEdBox txtCampo2 
         Height          =   285
         Left            =   9240
         TabIndex        =   2
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCampo1 
         Height          =   285
         Left            =   6990
         TabIndex        =   1
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblcampo2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final "
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   8400
         TabIndex        =   14
         Top             =   240
         Width           =   870
      End
      Begin VB.Label lblcampo1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5880
         TabIndex        =   13
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label2 
         Caption         =   "Seleccione Tipo de Cuota :"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmCuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim temporal
Dim loc_tipo As Integer
Dim PSCUO_LLAVE As rdoQuery
Dim cuo_rep01 As rdoResultset
Dim PSCUO_VENDEDOR As rdoQuery
Dim cuo_vendedor As rdoResultset



Private Sub Cmbtipos_Click()
loc_tipo = Val(Left(Cmbtipos.Text, 2))
End Sub

Private Sub Cmdclose_Click()
Unload frmCuotas
End Sub

Private Sub cmdmostrar_Click()
Dim PSCUO_VENDEDOR As rdoQuery
Dim cuo_vendedor As rdoResultset
pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = ? ORDER BY VEM_CODVEN"
Set PSCUO_VENDEDOR = CN.CreateQuery("", pub_cadena)
PSCUO_VENDEDOR(0) = LK_CODCIA
Set cuo_vendedor = PSCUO_VENDEDOR.OpenResultset(rdOpenKeyset, rdConcurValues)
    

cabe
Select Case loc_tipo
Case 1
    If Val(txt_key.Text) = 0 Then
      pub_cadena = "SELECT * FROM CUOTAS WHERE CUO_CODCIA = ? AND CUO_TIPO = ? AND CUO_FECHA1 = ? AND CUO_FECHA2 = ? "
    Else
      pub_cadena = "SELECT * FROM CUOTAS WHERE CUO_CODCIA = ? AND CUO_TIPO = ? AND CUO_FECHA1 = ? AND CUO_FECHA2 = ? " & _
      "AND CUO_CODVEN = ? "
    End If
    Set PSCUO_LLAVE = CN.CreateQuery("", pub_cadena)
    PSCUO_LLAVE(0) = LK_CODCIA
    PSCUO_LLAVE(1) = loc_tipo
    PSCUO_LLAVE(2) = txtCampo1.Text
    PSCUO_LLAVE(3) = txtCampo1.Text
    If Val(txt_key.Text) <> 0 Then PSCUO_LLAVE(4) = 0
    Set cuo_rep01 = PSCUO_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
    If cuo_rep01.EOF Then
        MsgBox "No existe Datos"
        Do Until cuo_vendedor.EOF
         gridigv.Rows = gridigv.Rows + 1
         gridigv.TextMatrix(gridigv.Rows - 1, 0) = cuo_vendedor!VEM_CODVEN
         gridigv.TextMatrix(gridigv.Rows - 1, 1) = cuo_vendedor!VEM_NOMBRE
         cuo_vendedor.MoveNext
        Loop
        Exit Sub
     End If

Case 2
Case 3
Case 4
Case 5
Case 6

End Select
'pub_cadena = "SELECT * FROM CUOTAS WHERE CUO_CODCIA = ? AND CUO_TIPO = ? AND CUO_FECHA1 = ? AND CUO_FECHA2 = ? " & _
'    "CUO_CODART = ? AND CUO_DIVISION = ? AND CUO_CODVEN = ? "

End Sub

Private Sub Form_Load()
Cmbtipos.AddItem "01 - Cuota por Vendedores"
Cmbtipos.AddItem "02 - Cuota por Vendedores y Articulos"
Cmbtipos.AddItem "03 - Cuota por Vendedores y Divisiones"
Cmbtipos.AddItem "04 - Cuota por Divisiones"
Cmbtipos.AddItem "05 - Cuota por Articulos"
Cmbtipos.AddItem "06 - Cuota por Empresa"
LLENA_GRUPOS cmbdivi, 122
txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
txtCampo2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")

End Sub

Private Sub i_nomarti_Click()

End Sub

Public Sub LLENA_GRUPOS(cont As ComboBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!tab_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
    
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
    TEXTOVAR.Text = gridigv.TextMatrix(gridigv.Row, gridigv.COL)
    TEXTOVAR.Visible = True
    Azul3 TEXTOVAR, TEXTOVAR
    TEXTOVAR.SetFocus
'End If
End Sub

Private Sub gridiGV_KeyUp(KeyCode As Integer, Shift As Integer)
Dim WC
Dim a, WF As Integer
Dim tf, t, tC
Dim SALE As Boolean
Dim Wsec

'If WMODO = "C" Then Exit Sub

'If cop_llave!COP_FLAG_MAYORIZACION = "M" Then
 'MsgBox "Ojo estaba Mayorizado..."
'End If


If Left(gridigv.TextMatrix(gridigv.Row, 0), 2) <> "MA" Then Exit Sub
 If KeyCode = 32 Then
  'If WMODO <> "C" Then Exit Sub
  tC = gridigv.COL
  For fila = 1 To gridigv.Cols - 1
      gridigv.COL = fila
      If gridigv.CellBackColor = QBColor(12) Then
         gridigv.CellBackColor = QBColor(15)
         gridigv.TextMatrix(gridigv.Row, 9) = "9"
      Else
         gridigv.CellBackColor = QBColor(12)
         gridigv.TextMatrix(gridigv.Row, 9) = "-1"
      End If
  Next fila
  gridigv.COL = tC
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
    gridigv.COL = 1
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
gridigv.Text = Format(TEXTOVAR.Text, "0.0000")
End Sub

Private Sub TEXTOVAR_GotFocus()
 temporal = gridigv.TextMatrix(gridigv.Row, gridigv.COL)
End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  TEXTOVAR.Text = temporal
  TEXTOVAR.Visible = False
  gridigv.SetFocus
  Exit Sub
End If
If gridigv.COL = 1 Then Consistencias gridigv, TEXTOVAR, KeyAscii
If gridigv.COL = 4 Then Consistencias gridigv, TEXTOVAR, KeyAscii
If gridigv.COL = 5 Or gridigv.COL = 6 Then Consistencias gridigv, TEXTOVAR, KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If
If gridigv.COL = 1 Or gridigv.COL = 4 Then
  If Val(TEXTOVAR.Text) > 99 Then
    Azul3 TEXTOVAR, TEXTOVAR
    Exit Sub
  End If
End If

'PUB_CAL_INI = gridigv.TextMatrix(gridigv.Row, 2)
'PUB_CAL_FIN = gridigv.TextMatrix(gridigv.Row, 2)
'pu_codcia = LK_CODCIA
'SQ_OPER = 1
'PUB_CODCIA = LK_CODCIA
'LEER_CAL_LLAVE
'cal_llave.Edit
'If gridigv.COL = 4 Then
'   cal_llave!cal_tipo_cambio = Val(TEXTOVAR.Text)
'End If
'If gridigv.COL = 1 Then
'   cal_llave!CAL_TC_MERCA = Val(TEXTOVAR.Text)
'   If Format(LK_FECHA_DIA, "dd/mm/yyyy") = Format(gridigv.TextMatrix(gridigv.Row, 0), "dd/mm/yyyy") Then
'      LK_TIPO_CAMBIO = Val(TEXTOVAR.Text)
'      'MDIForm1.StatusBar1.Panels(3).Text = "T.C.= S/. " + Format(LK_TIPO_CAMBIO, "0.0000")
'   End If
'End If
'If gridigv.COL = 5 Then
'   cal_llave!cal_tc_ingre = Val(TEXTOVAR.Text)
'End If
'If gridigv.COL = 6 Then
'   cal_llave!cal_tc_salid = Val(TEXTOVAR.Text)
'End If
'
'cal_llave.Update
'If gridigv.Row >= gridigv.Rows - 1 Then
'Else
'  gridigv.Row = gridigv.Row + 1
'End If
gridigv.SetFocus
TEXTOVAR.Visible = False

fin:

End Sub

Public Sub cabe()
gridigv.Clear
gridigv.Cols = 7
gridigv.Rows = 2
gridigv.ColWidth(0) = 800
gridigv.ColWidth(1) = 2500
gridigv.ColWidth(2) = 1000
gridigv.ColWidth(3) = 1000
gridigv.ColWidth(4) = 1000
gridigv.ColWidth(5) = 1000
gridigv.ColWidth(6) = 1000

gridigv.TextMatrix(0, 0) = "Codigo"
gridigv.TextMatrix(1, 0) = ""
gridigv.TextMatrix(0, 1) = "Descripcion"
gridigv.TextMatrix(1, 1) = ""

gridigv.TextMatrix(0, 2) = "  "
gridigv.TextMatrix(1, 2) = "U.M"
gridigv.TextMatrix(0, 3) = "Cuota "
gridigv.TextMatrix(1, 3) = "Cantidad"

gridigv.TextMatrix(0, 4) = "Avance "
gridigv.TextMatrix(1, 4) = "Unidades"
gridigv.TextMatrix(0, 5) = "Cuota"
gridigv.TextMatrix(1, 5) = "Valor"
gridigv.TextMatrix(0, 6) = "Avance"
gridigv.TextMatrix(1, 6) = "Valor"

End Sub

Private Sub Consistencias(wsGrid As MSFlexGrid, wsTexto As RichTextBox, wsKeyAscii As Integer)
  Static VALOR
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

