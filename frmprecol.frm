VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmprecol 
   Caption         =   "Definición de Precios"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox textovar 
      Height          =   285
      Left            =   3360
      TabIndex        =   5
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdgrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "Ce&rrar"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.ComboBox colores 
      Height          =   315
      Left            =   840
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin MSFlexGridLib.MSFlexGrid gridpro 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6165
      _Version        =   327680
   End
   Begin VB.Label lblcolor 
      Caption         =   "Color :"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmprecol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temporal
Dim PS_PROV As rdoQuery
Dim llave_prov As rdoResultset
Dim LOC_TIPREG As Integer
Dim PS_PRECIO As rdoQuery
Dim llave_prec As rdoResultset


Private Sub cmdcerrar_Click()
Unload frmprecol
End Sub

Private Sub cmdgrabar_Click()
Dim WNUMTAB  As Integer
pub_cadena = "DELETE TABLAS WHERE TAB_CODCIA = '" & "00" & "' AND TAB_CODART = " & Trim(Right(colores.Text, 6)) & " AND TAB_TIPREG = " & LOC_TIPREG
CN.Execute pub_cadena, rdExecDirect
SQ_OPER = 2
PUB_TIPREG = LOC_TIPREG
PUB_CODCIA = "00"
LEER_TAB_LLAVE
tab_mayor.MoveLast
If tab_mayor.EOF Then
  WNUMTAB = 1
Else
  WNUMTAB = tab_mayor!TAB_NUMTAB
End If
For fila = 1 To gridpro.Rows - 1
    tab_mayor.AddNew
    tab_mayor!TAB_CODCIA = "00"
    tab_mayor!TAB_TIPREG = LOC_TIPREG
    WNUMTAB = WNUMTAB + 1
    tab_mayor!TAB_NUMTAB = WNUMTAB
    tab_mayor!TAB_NOMLARGO = Val(gridpro.TextMatrix(fila, 2))
    tab_mayor!TAB_NOMCORTO = ""
    tab_mayor!TAB_CONTABLE2 = Trim(gridpro.TextMatrix(fila, 0))
    tab_mayor!TAB_CODART = Val(Right(colores.Text, 6))
    tab_mayor!TAB_CODCLIE = 0 'Val(gridpro.TextMatrix(fila, 0))
    tab_mayor.Update

Next fila
cmdgrabar.Enabled = False
End Sub

Private Sub colores_Click()
cmdgrabar.Enabled = False
PS_PROV(0) = LK_CODCIA
PS_PROV(1) = "P"
llave_prov.Requery
gridpro.Clear
gridpro.Cols = 6
gridpro.ColWidth(0) = 600
gridpro.ColWidth(1) = 2000
gridpro.ColWidth(2) = 900
gridpro.ColWidth(3) = 0
gridpro.ColWidth(4) = 0
gridpro.ColWidth(5) = 0
gridpro.TextMatrix(0, 0) = "Cod."
gridpro.TextMatrix(0, 1) = "Proveedor"
gridpro.TextMatrix(0, 2) = "Precio"
gridpro.Rows = 1
fila = 1
PS_PRECIO(0) = "00"
Do Until llave_prov.EOF
   gridpro.Rows = gridpro.Rows + 1
   gridpro.TextMatrix(fila, 0) = llave_prov!CLI_RUC_ESPOSO
   gridpro.TextMatrix(fila, 1) = llave_prov!CLI_NOMBRE
   PS_PRECIO(1) = Val(Right(colores.Text, 6))
   PS_PRECIO(2) = llave_prov!CLI_RUC_ESPOSO
   llave_prec.Requery
   If llave_prec.EOF Then
    gridpro.TextMatrix(fila, 2) = 0
   Else
    gridpro.TextMatrix(fila, 2) = Format(Val(llave_prec!TAB_NOMLARGO), "0.00")
   End If
   fila = fila + 1
llave_prov.MoveNext
Loop
End Sub

Private Sub colores_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If gridpro.Rows <> 1 Then
    gridpro.COL = 2
    gridpro.Row = 1
  End If
  gridpro.SetFocus
  Exit Sub
End If

End Sub

Private Sub Form_Load()
CenterMe frmprecol
If OP_FORM = "C" Then
  LOC_TIPREG = 232
Else
  LOC_TIPREG = 230
End If
pub_cadena = "SELECT CLI_CODCLIE, CLI_NOMBRE, CLI_RUC_ESPOSO FROM CLIENTES  WHERE CLI_CODCIA = ? AND CLI_CP = ? AND CLI_GRUPO = 1 ORDER BY CLI_NOMBRE"
Set PS_PROV = CN.CreateQuery("", pub_cadena)
PS_PROV(0) = 0
PS_PROV(1) = 0
Set llave_prov = PS_PROV.OpenResultset(rdOpenKeyset, rdConcurValues)


pub_cadena = "SELECT * FROM TABLAS WHERE TAB_TIPREG = " & LOC_TIPREG & " AND TAB_CODCIA = ? AND TAB_CODART = ? AND TAB_CONTABLE2 = ? "
Set PS_PRECIO = CN.CreateQuery("", pub_cadena)
PS_PRECIO(0) = 0
PS_PRECIO(1) = 0
PS_PRECIO(2) = 0
Set llave_prec = PS_PRECIO.OpenResultset(rdOpenKeyset, rdConcurValues)

SQ_OPER = 2
If LOC_TIPREG = 232 Then
  PUB_TIPREG = 132
Else
  PUB_TIPREG = 130
End If
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
colores.Clear
Do Until tab_mayor.EOF
  colores.AddItem tab_mayor!TAB_NOMLARGO & String(80, " ") & tab_mayor!TAB_NUMTAB
tab_mayor.MoveNext
Loop
cmdgrabar.Enabled = False

End Sub
Private Sub gridpro_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, WC
Static CONS
If KeyAscii = 27 Then
  colores.SetFocus
  Exit Sub
End If
If KeyAscii <> 13 Then Exit Sub

'If Trim(gridpro.TextMatrix(gridpro.Row, 9)) <> "8" Then
'  If Trim(gridpro.TextMatrix(gridpro.Row, 0)) = "" Then Exit Sub
'  If Trim(gridpro.TextMatrix(gridpro.Row, 1)) <> "" And gridpro.Col = 2 Or gridpro.Col = 3 Then GoTo leer
'  If Trim(gridpro.TextMatrix(gridpro.Row, 8)) <> "0" Then Exit Sub
'End If


If gridpro.Row = 0 Then Exit Sub
If gridpro.COL = 1 Then Exit Sub
'   a = Val(gridpro.TextMatrix(gridpro.Row - 1, 0))
'   a = a + 1
'  gridpro.TextMatrix(gridpro.Row, 0) = a
'End If
'If WMODO = "I" Or WMODO = "C" Then
    textovar.Left = gridpro.Left + gridpro.CellLeft
    textovar.Width = gridpro.CellWidth
    textovar.Height = gridpro.CellHeight
    textovar.Top = gridpro.Top + gridpro.CellTop
    textovar.Text = gridpro.TextMatrix(gridpro.Row, gridpro.COL)
    textovar.Visible = True
    Azul textovar, textovar
    textovar.SetFocus
'End If
End Sub

Private Sub gridpro_KeyUp(KeyCode As Integer, Shift As Integer)
Dim WC
Dim a, WF As Integer
Dim tf, t, TC
Dim SALE As Boolean
Dim Wsec

'If WMODO = "C" Then Exit Sub

'If cop_llave!COP_FLAG_MAYORIZACION = "M" Then
 'MsgBox "Ojo estaba Mayorizado..."
'End If


If Left(gridpro.TextMatrix(gridpro.Row, 0), 2) <> "MA" Then Exit Sub
 If KeyCode = 32 Then
  'If WMODO <> "C" Then Exit Sub
  TC = gridpro.COL
  For fila = 1 To gridpro.Cols - 1
      gridpro.COL = fila
      If gridpro.CellBackColor = QBColor(12) Then
         gridpro.CellBackColor = QBColor(15)
         gridpro.TextMatrix(gridpro.Row, 9) = "9"
      Else
         gridpro.CellBackColor = QBColor(12)
         gridpro.TextMatrix(gridpro.Row, 9) = "-1"
      End If
  Next fila
  gridpro.COL = TC
  gridpro.SetFocus
  Exit Sub
End If
If KeyCode = 45 Then
    Wsec = Wsec + 1
    If Trim(gridpro.TextMatrix(gridpro.Row + 1, 11)) = "8" Then
         Exit Sub
    Else
      If Trim(gridpro.TextMatrix(gridpro.Row + 1, 0)) = "T" Then Exit Sub
    End If
    If Val(gridpro.TextMatrix(gridpro.Row, 4)) = 0 And Val(gridpro.TextMatrix(gridpro.Row, 5)) = 0 Then Exit Sub
    gridpro.AddItem "", gridpro.Row + 1
    gridpro.TextMatrix(gridpro.Row + 1, 0) = "MAN. " & Format(gridpro.TextMatrix(gridpro.Row, 10), "dd/mm/yyyy")
    gridpro.TextMatrix(gridpro.Row + 1, 6) = Wsec
    gridpro.TextMatrix(gridpro.Row + 1, 8) = gridpro.TextMatrix(gridpro.Row, 8)
    gridpro.TextMatrix(gridpro.Row + 1, 3) = gridpro.TextMatrix(gridpro.Row, 3)
    gridpro.TextMatrix(gridpro.Row + 1, 7) = gridpro.TextMatrix(gridpro.Row, 7)
    gridpro.TextMatrix(gridpro.Row + 1, 10) = gridpro.TextMatrix(gridpro.Row, 10)
    gridpro.TextMatrix(gridpro.Row + 1, 11) = "8"
    gridpro.Row = gridpro.Row + 1
    gridpro.COL = 1
    gridpro.SetFocus
End If
Exit Sub
If KeyCode = 46 Then
If gridpro.Rows <= 3 Then
Else
   pub_mensaje = MsgBox("Desea Quitar el Item de la Cuenta : " & Trim(gridpro.TextMatrix(gridpro.Row, 1)), vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
   If pub_mensaje = vbNo Then
     gridpro.SetFocus
     Exit Sub
   Else
     gridpro.RowHeight(gridpro.Row) = 1
     gridpro.Row = gridpro.Row + 1
    
   'gridpro.RemoveItem (gridpro.Row)
   'gridpro.Refresh
   gridpro.SetFocus
   End If
End If
End If
'gridpro.SetFocus
Exit Sub



End Sub

Private Sub gridpro_Scroll()
textovar.Visible = False
End Sub


Private Sub textovar_Change()
gridpro.Text = Format(textovar.Text, "0.00")
End Sub

Private Sub TEXTOVAR_GotFocus()
 temporal = gridpro.TextMatrix(gridpro.Row, gridpro.COL)
End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  textovar.Text = temporal
  textovar.Visible = False
  gridpro.SetFocus
  Exit Sub
End If
If gridpro.COL = 3 Then Consistencias gridpro, textovar, KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If
If gridpro.COL = 1 Or gridpro.COL = 4 Then
  If Val(textovar.Text) > 99 Then
    Azul3 textovar, textovar
    Exit Sub
  End If
End If
If gridpro.Row >= gridpro.Rows - 1 Then
Else
  gridpro.Row = gridpro.Row + 1
End If
gridpro.SetFocus
textovar.Visible = False
cmdgrabar.Enabled = True
cmdgrabar_Click
fin:

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

