VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmPlan 
   BackColor       =   &H00FAEFDA&
   Caption         =   "Definición de Voucher"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   Icon            =   "FrmPlan.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdcerrar 
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
      Height          =   360
      Left            =   9660
      TabIndex        =   0
      Top             =   6195
      Width           =   1095
   End
   Begin VB.Frame fravoucher 
      BackColor       =   &H00FAEFDA&
      Height          =   6600
      Left            =   915
      TabIndex        =   1
      Top             =   255
      Width           =   10290
      Begin RichTextLib.RichTextBox TEXTOVAR 
         Height          =   375
         Left            =   6315
         TabIndex        =   5
         Top             =   2265
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393217
         BackColor       =   16445402
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"FrmPlan.frx":0442
      End
      Begin MSFlexGridLib.MSFlexGrid grid_fac 
         Height          =   2385
         Left            =   3600
         TabIndex        =   11
         Tag             =   "9999"
         Top             =   1275
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   4207
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   8388608
         ForeColorFixed  =   16777215
         BackColorSel    =   15386515
         GridColorFixed  =   16445402
         FocusRect       =   2
         HighLight       =   2
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
      Begin VB.CommandButton Command1 
         Caption         =   "Copiar a otra Compañia. !!!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   495
         TabIndex        =   20
         Top             =   5895
         Width           =   2295
      End
      Begin VB.ListBox listv 
         Height          =   3570
         Left            =   180
         TabIndex        =   2
         Top             =   1185
         Width           =   3135
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FAEFDA&
         Height          =   855
         Left            =   150
         TabIndex        =   17
         Top             =   4800
         Width           =   3240
         Begin VB.CommandButton cmdnuevo 
            Caption         =   "Nuevo Vouc."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   300
            Width           =   1455
         End
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "Eliminar Vouc."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   18
            Top             =   300
            Width           =   1215
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grid_des 
         Height          =   2280
         Left            =   3630
         TabIndex        =   16
         Tag             =   "9999"
         Top             =   4215
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   4022
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   8388608
         ForeColorFixed  =   16777215
         BackColorSel    =   15386515
         GridColorFixed  =   12961418
         FocusRect       =   2
         HighLight       =   2
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
      Begin VB.ComboBox proce 
         BackColor       =   &H00FFFFFF&
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
         ItemData        =   "FrmPlan.frx":04C4
         Left            =   165
         List            =   "FrmPlan.frx":04C6
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   420
         Width           =   3135
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FAEFDA&
         Height          =   2730
         Left            =   3480
         TabIndex        =   10
         Top             =   1065
         Width           =   6525
         Begin VB.CommandButton cmdeliminaes 
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
            Height          =   375
            Left            =   5190
            TabIndex        =   24
            Top             =   1830
            Width           =   1095
         End
         Begin VB.CommandButton cmdnuevoes 
            Caption         =   "&Nuevo"
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
            Left            =   5190
            TabIndex        =   23
            Top             =   450
            Width           =   1095
         End
         Begin VB.CommandButton cmdgrabar 
            Caption         =   "&Grabar"
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
            Left            =   5190
            TabIndex        =   22
            Top             =   1140
            Width           =   1095
         End
      End
      Begin VB.TextBox tglosa 
         Height          =   285
         Left            =   4080
         TabIndex        =   8
         Top             =   720
         Width           =   3975
      End
      Begin ComctlLib.ProgressBar PB 
         Height          =   165
         Left            =   240
         TabIndex        =   6
         Top             =   6330
         Visible         =   0   'False
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   291
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.Label lblpla 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destino de Cuentas Analiticas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   195
         Index           =   3
         Left            =   3645
         TabIndex        =   15
         Top             =   3870
         Width           =   2505
      End
      Begin VB.Label lblpla 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Movimiento:"
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
         Index           =   4
         Left            =   195
         TabIndex        =   14
         Top             =   150
         Width           =   1710
      End
      Begin VB.Label lcodigo 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   5640
         TabIndex        =   12
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblpla 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Glosa :"
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
         Index           =   2
         Left            =   3480
         TabIndex        =   9
         Top             =   735
         Width           =   555
      End
      Begin VB.Label lvoucher 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label lblpla 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Definición de Estructura:"
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
         Index           =   1
         Left            =   3480
         TabIndex        =   4
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label lblpla 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Voucher"
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
         Index           =   0
         Left            =   195
         TabIndex        =   3
         Top             =   870
         Width           =   1395
      End
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
      Left            =   30
      TabIndex        =   21
      Top             =   6915
      Width           =   11895
   End
End
Attribute VB_Name = "FrmPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim plt_llave As rdoResultset
Dim PSPL_LLAVE As rdoQuery
Dim plt_mayor As rdoResultset
Dim PSPL_MAYOR As rdoQuery
Dim plt_mayor2 As rdoResultset
Dim PSPL_MAYOR2 As rdoQuery
Dim loc_key As Integer

Private Sub cmdcerrar_Click()
Unload FrmPlan
End Sub

Private Sub cmdeliminaes_Click()
 If grid_fac.Rows <= 2 Then Exit Sub
 grid_fac.RemoveItem grid_fac.Row
 FrmPlan.cmdgrabar.Enabled = True
End Sub

Private Sub cmdeliminar_Click()
Dim WNUM As Integer
If Trim(listv.Text) = "" Then
  MsgBox "Seleccione uno de la lista.", 48, Pub_Titulo
  listv.SetFocus
  Exit Sub
End If
pub_mensaje = "Esta seguro que desea Eliminar este Voucher...?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If
WNUM = Val(Trim(Right(listv.Text, 6)))
pub_cadena = "DELETE PLANTILLA  WHERE PLT_CODCIA = '" & LK_CODCIA & "' AND PLT_NUMERO = " & WNUM
CN.Execute pub_cadena, rdExecDirect
listv.RemoveItem listv.ListIndex
cabegrid
tglosa.Text = ""
lvoucher.Caption = ""
lcodigo.Caption = ""
listv.SetFocus

End Sub

Private Sub cmdgrabar_Click()
Dim wssecu As Integer
Dim wglosa
If Trim(listv.Text) = "" Then
   cmdgrabar.Enabled = False
  Exit Sub
End If
 
For fila = 1 To grid_fac.Rows - 1
If Trim(grid_fac.TextMatrix(fila, 0)) <> "" Then
  If Trim(grid_fac.TextMatrix(fila, 2)) = "" Then
 '  MsgBox "Verificar su Debe o Haber.", 48, Pub_Titulo
 '  grid_fac.SetFocus
 '  Exit Sub
  End If
End If
Next fila
wssecu = 0
WNUM = Val(Trim(Right(listv.Text, 6)))
pub_cadena = "DELETE PLANTILLA  WHERE PLT_CODCIA = '" & LK_CODCIA & "' AND PLT_NUMERO = " & WNUM & " AND PLT_TIPMOV = " & Trim(Left(proce.Text, 3))

If UCase(Left(listv.Text, 3)) = "ING" And Val(Left(proce.Text, 3)) = 3 Then WNUM = 100
If UCase(Left(listv.Text, 3)) = "EGR" And Val(Left(proce.Text, 3)) = 3 Then WNUM = 126


CN.Execute pub_cadena, rdExecDirect

plt_llave.AddNew
plt_llave!PLT_CODCIA = LK_CODCIA
plt_llave!PLT_NUMERO = WNUM
plt_llave!PLT_SECUENCIA = wssecu
plt_llave!PLT_NOMBRE = Left(listv.Text, 40)
plt_llave!PLT_CUENTA = " "
plt_llave!PLT_DH = " "
plt_llave!PLT_GLOSA = Trim(tglosa.Text)
plt_llave!PLT_tipmov = Val(Left(proce.Text, 3))
cmdgrabar.Enabled = True
plt_llave.Update
For fila = 1 To grid_fac.Rows - 1
    If Val(grid_fac.TextMatrix(fila, 0)) <= 0 Then
      GoTo pasa
    End If
    wssecu = wssecu + 1
    plt_llave.AddNew
    plt_llave!PLT_CODCIA = LK_CODCIA
    plt_llave!PLT_NUMERO = WNUM
    plt_llave!PLT_SECUENCIA = wssecu
    plt_llave!PLT_CUENTA = grid_fac.TextMatrix(fila, 0)
    plt_llave!PLT_DH = grid_fac.TextMatrix(fila, 2)
    'If Trim(grid_fac.TextMatrix(fila, 4)) = "X" Then
      plt_llave!PLT_NOMBRE = grid_fac.TextMatrix(fila, 4)
    'Else
    '  plt_llave!PLT_NOMBRE = ""
    'End If
    plt_llave!PLT_GLOSA = Trim(tglosa.Text)
    plt_llave!PLT_tipmov = Val(Left(proce.Text, 3))
    plt_llave.Update
Next fila
pasa:
MsgBox "Datos Grabados.", 48, Pub_Titulo
cmdgrabar.Enabled = False
listv.SetFocus

End Sub

Private Sub cmdnuevo_Click()
Dim www As String
www = InputBox("Ingrese la Descripción del Voucher :", "Voucher", "Nuevo...")
If www = "" Then Exit Sub
PSPL_MAYOR2(0) = LK_CODCIA
plt_mayor2.Requery
If plt_mayor2.EOF Then
 WNUM = 1
Else
 WNUM = Val(plt_mayor2!PLT_NUMERO) + 1
End If

plt_llave.AddNew
plt_llave!PLT_CODCIA = LK_CODCIA
plt_llave!PLT_NUMERO = WNUM
plt_llave!PLT_SECUENCIA = 0
plt_llave!PLT_NOMBRE = www
plt_llave!PLT_CUENTA = " "
plt_llave!PLT_DH = ""
plt_llave!PLT_GLOSA = Trim(tglosa.Text)
plt_llave!PLT_tipmov = Val(Left(proce.Text, 3))
plt_llave.Update
listv.AddItem www & String(80, " ") & WNUM
listv.SetFocus

End Sub


Private Sub cmdnuevoes_Click()
If Trim(listv.Text) = "" Then
  MsgBox "Seleccione uno de la lista.", 48, Pub_Titulo
  listv.SetFocus
  Exit Sub
End If
grid_fac.Rows = grid_fac.Rows + 1
grid_fac.TextMatrix(grid_fac.Rows - 1, 0) = " "
grid_fac.TextMatrix(grid_fac.Rows - 1, 1) = " "
grid_fac.TextMatrix(grid_fac.Rows - 1, 2) = " "
grid_fac.TextMatrix(grid_fac.Rows - 1, 3) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 4) = ""
grid_fac.Row = grid_fac.Rows - 1
grid_fac.Col = 2
grid_fac.CellAlignment = 4
grid_fac.Col = 0
grid_fac.SetFocus
End Sub

Private Sub Command1_Click()
On Error GoTo SALE
Dim PSPL_LLAVE10 As rdoQuery
Dim plt_llave10 As rdoResultset

Dim PSPL_OTRA As rdoQuery
Dim plt_otra As rdoResultset
Dim WS As String * 2
WS = InputBox("A QUE COMPRAÑIA DESEA COPIAR LA PLANTILLA (Ejemplo: 02 o 03 o 04...) : ", "COPIAR", "")
If Trim(WS) = "" Then Exit Sub

pub_cadena = "SELECT * FROM PLANTILLA WHERE PLT_CODCIA = ? ORDER BY PLT_CODCIA "
Set PSPL_OTRA = CN.CreateQuery("", pub_cadena)
PSPL_OTRA(0) = 0
Set plt_otra = PSPL_OTRA.OpenResultset(rdOpenKeyset, rdConcurValues)
PSPL_OTRA(0) = WS
plt_otra.Requery


pub_cadena = "SELECT * FROM PLANTILLA WHERE PLT_CODCIA = ? ORDER BY PLT_CODCIA "
Set PSPL_LLAVE10 = CN.CreateQuery("", pub_cadena)
PSPL_LLAVE10(0) = 0
Set plt_llave10 = PSPL_LLAVE10.OpenResultset(rdOpenKeyset, rdConcurValues)
PSPL_LLAVE10(0) = LK_CODCIA
plt_llave10.Requery

Do Until plt_llave10.EOF
    plt_otra.AddNew
    For fila = 0 To plt_llave10.rdoColumns.Count - 1
      If fila = 0 Then
        plt_otra(fila) = WS
      Else
        plt_otra(fila) = plt_llave10(fila)
      End If
    Next fila
    plt_otra.Update
plt_llave10.MoveNext
Loop
MsgBox "Copiado Terminado", 48, Pub_Titulo

Exit Sub
SALE:
MsgBox "Existen Datos en Otra Compañia, Verificar", 48, Pub_Titulo
End Sub

Private Sub Form_Activate()
proce.SetFocus
End Sub

Private Sub Form_Load()
CenterMe FrmPlan
pub_cadena = "SELECT PLT_NUMERO FROM PLANTILLA WHERE PLT_CODCIA = ?  ORDER BY PLT_NUMERO DESC"
Set PSPL_MAYOR2 = CN.CreateQuery("", pub_cadena)
PSPL_MAYOR2(0) = 0
PSPL_MAYOR2.MaxRows = 1
Set plt_mayor2 = PSPL_MAYOR2.OpenResultset(rdOpenKeyset, rdConcurValues)


pub_cadena = "SELECT * FROM PLANTILLA WHERE PLT_CODCIA = ? AND PLT_NUMERO = ?  ORDER BY PLT_SECUENCIA "
Set PSPL_LLAVE = CN.CreateQuery("", pub_cadena)
PSPL_LLAVE(0) = 0
PSPL_LLAVE(1) = 0
Set plt_llave = PSPL_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM PLANTILLA WHERE PLT_CODCIA = ?  AND PLT_TIPMOV = ? AND PLT_SECUENCIA = 0  ORDER BY PLT_NOMBRE "
Set PSPL_MAYOR = CN.CreateQuery("", pub_cadena)
PSPL_MAYOR(0) = 0
PSPL_MAYOR(1) = 0
Set plt_mayor = PSPL_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
cmdgrabar.Enabled = False


proce.Clear
PUB_TIPREG = 150
PUB_CODCIA = "00"
SQ_OPER = 2
LEER_TAB_LLAVE
Do Until tab_mayor.EOF
    proce.AddItem "0" & tab_mayor!TAB_NUMTAB & ".-" & Trim(tab_mayor!tab_nomlargo) & String(80, " ") & Trim(tab_mayor!TAB_CONTABLE2)
    proce.ItemData(proce.NewIndex) = tab_mayor!TAB_NUMTAB
    tab_mayor.MoveNext
Loop

End Sub

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub grid_fac_DblClick()
If grid_fac.Col <> 4 Then Exit Sub
If Trim(grid_fac.TextMatrix(grid_fac.Row, 4)) = "X" Then
   grid_fac.TextMatrix(grid_fac.Row, 4) = ""
Else
  grid_fac.TextMatrix(grid_fac.Row, 4) = "X"
End If
cmdgrabar.Enabled = True
End Sub

Private Sub grid_fac_EnterCell()
llena_destinos grid_fac.TextMatrix(grid_fac.Row, 0)
End Sub

Private Sub grid_fac_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, WC
Static CONS
If KeyAscii <> 13 Then Exit Sub
If Trim(listv.Text) = "" Then Exit Sub
If grid_fac.Col = 1 Then Exit Sub
If grid_fac.Col = 4 Then
 ' grid_fac_DblClick
  'Exit Sub
End If

If grid_fac.Col > 4 Then
  Exit Sub
End If
'If grid_fac.Col = 4 Then
'  If Val(grid_fac.TextMatrix(grid_fac.Row, 0)) = 0 And WMODO = "I" Then Exit Sub
'  If Val(grid_fac.TextMatrix(grid_fac.Row, 1)) = 0 And WMODO = "C" Then Exit Sub
'  If Val(grid_fac.TextMatrix(grid_fac.Row, 4)) = 0 And Val(grid_fac.TextMatrix(grid_fac.Row, 5)) = 0 Then
'  Else
'    If Val(grid_fac.TextMatrix(grid_fac.Row, 4)) = 0 Then Exit Sub
'  End If
'End If''

'If grid_fac.Col = 5 Then
'  If Val(grid_fac.TextMatrix(grid_fac.Row, 0)) = 0 And WMODO = "I" Then Exit Sub
'  If Val(grid_fac.TextMatrix(grid_fac.Row, 1)) = 0 And WMODO = "C" Then Exit Sub
'  If Val(grid_fac.TextMatrix(grid_fac.Row, 4)) = 0 And Val(grid_fac.TextMatrix(grid_fac.Row, 5)) = 0 Then
'  Else
'     If Val(grid_fac.TextMatrix(grid_fac.Row, 5)) = 0 Then Exit Sub
'  End If
'End If

    TEXTOVAR.Left = grid_fac.Left + grid_fac.CellLeft 'fravoucher.Left +
    TEXTOVAR.Width = grid_fac.CellWidth
    TEXTOVAR.Height = grid_fac.CellHeight
    TEXTOVAR.Top = grid_fac.Top + grid_fac.CellTop  ' - 120 '480fravoucher.Top +
    TEXTOVAR.Text = grid_fac.TextMatrix(grid_fac.Row, grid_fac.Col)
    wfila_act = grid_fac.Row
    TEXTOVAR.Visible = True
    Azul3 TEXTOVAR, TEXTOVAR
    TEXTOVAR.SetFocus
    
Exit Sub
leer:
If grid_fac.TextMatrix(grid_fac.Row, 1) <> "" Then
 SQ_OPER = 1
 PUB_CUENTA = grid_fac.TextMatrix(grid_fac.Row, 1)
 PUB_CODCIA = LK_CODCIA
 LEER_COM_LLAVE
 If Not com_llave.EOF Then
   grid_fac.TextMatrix(grid_fac.Row, 2) = com_llave!com_DESCRIPCION
 End If
End If


End Sub

Private Sub grid_fac_KeyUp(KeyCode As Integer, Shift As Integer)
Dim WC
Dim a, WF As Integer
Dim tf, t, tc
Dim SALE As Boolean

'If Left(grid_fac.TextMatrix(grid_fac.Row, 0), 2) <> "MA" Then Exit Sub
If KeyCode = 45 Then
 If grid_fac.Row + 1 = grid_fac.Rows Then
   grid_fac.Rows = grid_fac.Rows + 1
 Else
   grid_fac.AddItem "", grid_fac.Row
 End If
  'cmdnuevoes_Click
  Exit Sub
End If
If KeyCode = 46 Then
  cmdeliminaes_Click
  Exit Sub
End If
 
If KeyCode = 45 Then
    If grid_fac.Row >= grid_fac.Rows - 1 Then Exit Sub
    Wsec = Wsec + 1
    If Trim(grid_fac.TextMatrix(grid_fac.Row + 1, 11)) = "8" Then
         Exit Sub
    Else
      If Trim(grid_fac.TextMatrix(grid_fac.Row + 1, 0)) = "T" Then Exit Sub
    End If
    If Val(grid_fac.TextMatrix(grid_fac.Row, 4)) = 0 And Val(grid_fac.TextMatrix(grid_fac.Row, 5)) = 0 Then Exit Sub
    grid_fac.AddItem "", grid_fac.Row + 1
    'grid_fac.TextMatrix(grid_fac.Row + 1, 0) = "MAN. " & Format(grid_fac.TextMatrix(grid_fac.Row, 10), "dd/mm/yyyy")
    grid_fac.TextMatrix(grid_fac.Row + 1, 0) = Format(grid_fac.TextMatrix(grid_fac.Row, 10), "dd/mm/yyyy")
    grid_fac.TextMatrix(grid_fac.Row + 1, 6) = Wsec
    grid_fac.TextMatrix(grid_fac.Row + 1, 8) = grid_fac.TextMatrix(grid_fac.Row, 8)
    grid_fac.TextMatrix(grid_fac.Row + 1, 3) = grid_fac.TextMatrix(grid_fac.Row, 3)
    grid_fac.TextMatrix(grid_fac.Row + 1, 7) = grid_fac.TextMatrix(grid_fac.Row, 7)
    grid_fac.TextMatrix(grid_fac.Row + 1, 10) = grid_fac.TextMatrix(grid_fac.Row, 10)
    grid_fac.TextMatrix(grid_fac.Row + 1, 11) = "8"
    grid_fac.Row = grid_fac.Row + 1
    grid_fac.Col = 1
    grid_fac.SetFocus
End If
'Exit Sub
'If KeyCode = 46 And Left(cmdIngreso.Caption, 2) = "&G" Then
'If grid_fac.Rows <= 3 Then
'Else
'   pub_mensaje = MsgBox("Desea Quitar el Item de la Cuenta : " & Trim(grid_fac.TextMatrix(grid_fac.Row, 1)), vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
'   If pub_mensaje = vbNo Then
'     grid_fac.SetFocus
'     Exit Sub
'   Else
'   grid_fac.RemoveItem (grid_fac.Row)
'   grid_fac.Refresh
'   grid_fac.SetFocus
'   End If
'End If
Exit Sub



End Sub



Private Sub listv_Click()
llena_data Val(Right(listv.Text, 6))
End Sub

Private Sub listv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul tglosa, tglosa
End If
End Sub

Private Sub listv_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
Dim www As String
Dim WI
If Trim(listv.Text) = "" Then Exit Sub
 www = InputBox("Modificar la Descripción del Voucher :", "Modificar Voucher", Trim(Left(listv.Text, 40)))
 If www = "" Then Exit Sub
 WI = listv.ListIndex
 WCOD = Trim(Right(listv.Text, 6))
 listv.RemoveItem WI
 listv.AddItem www & String(80, " ") & WCOD, WI
 listv.ListIndex = WI
 cmdgrabar_Click

End If
End Sub

Private Sub proce_Click()
listv.Clear
grid_fac.Clear
tglosa.Text = ""
lvoucher.Caption = ""
lcodigo.Caption = ""
cabegrid
PSPL_MAYOR(0) = LK_CODCIA
PSPL_MAYOR(1) = Val(Left(proce.Text, 3))
plt_mayor.Requery
Do Until plt_mayor.EOF
  listv.AddItem plt_mayor!PLT_NOMBRE & String(80, " ") & plt_mayor!PLT_NUMERO
  plt_mayor.MoveNext
Loop
cmdgrabar.Enabled = False
listv.SetFocus
End Sub

Private Sub textovar_Change()
TEXTOVAR.MaxLength = 0
If grid_fac.Col = 2 Then
  TEXTOVAR.MaxLength = 1
  If UCase(TEXTOVAR.Text) = "D" Or UCase(TEXTOVAR.Text) = "H" Then
    TEXTOVAR.Text = UCase(TEXTOVAR.Text)
  Else
    TEXTOVAR.Text = ""
  Exit Sub
  End If
End If
'If grid_fac.Col = 1 Then
'Else
 grid_fac.Text = TEXTOVAR.Text
 cmdgrabar.Enabled = True
' suma_grid
' suma_subtotal
'End If
End Sub

Private Sub TEXTOVAR_GotFocus()
temporal = grid_fac.TextMatrix(grid_fac.Row, grid_fac.Col)
End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  TEXTOVAR.Text = temporal
  TEXTOVAR.Visible = False
  grid_fac.SetFocus
  Exit Sub
End If
'If grid_fac.Col = 4 Or grid_fac.Col = 5 Then Consistencias grid_fac, TEXTOVAR, KeyAscii
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If grid_fac.Col = 0 Then
 If Chr(KeyAscii) <> "*" Then
  SOLO_ENTERO KeyAscii
 End If
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
If TEXTOVAR.Text = "" Then Exit Sub
If Left(TEXTOVAR.Text, 1) = "*" Then
  BUSCAR_CTA 0
  Exit Sub
End If
If grid_fac.Col = 0 Then
 If WMODO = "C" Then
   If Not IsDate(TEXTOVAR.Text) Then
      TEXTOVAR.SetFocus
      Exit Sub
   End If
 End If
End If



If grid_fac.Col = 0 Then

 If TEXTOVAR.Text = "" Then
    TEXTOVAR.Visible = False
    llave1 = ""
    grid_fac.SetFocus
    Exit Sub
 Else
    SQ_OPER = 1
    PUB_CUENTA = TEXTOVAR.Text
    PUB_CODCIA = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
         MsgBox "La Cuenta que Digito NO EXISTE  : " + Trim(TEXTOVAR.Text), 48, Pub_Titulo
         TEXTOVAR.Text = ""
         grid_fac.TextMatrix(grid_fac.Row, 0) = ""
         grid_fac.TextMatrix(grid_fac.Row, 1) = ""
         grid_fac.TextMatrix(grid_fac.Row, 2) = ""
         grid_fac.TextMatrix(grid_fac.Row, 3) = ""
         grid_fac.TextMatrix(grid_fac.Row, 4) = ""
         TEXTOVAR.SetFocus
      '   grid_fac.SetFocus
         GoTo fin
    Else
      'If Val(com_llave!COM_NIVEL) <> Val(cop_llave!cop_nivel_max) - 1 Then
      '    MsgBox "Cuenta es validad, pero no es Analitica ...", 48, Pub_Titulo
      '    Azul3 TEXTOVAR, TEXTOVAR
      '    GoTo fin
      'End If
     grid_fac.TextMatrix(grid_fac.Row, 0) = Trim(com_llave!com_cuenta)
     grid_fac.TextMatrix(grid_fac.Row, 1) = Trim(com_llave!com_DESCRIPCION)
    End If
  End If
End If


If grid_fac.Col = 0 Then
  TEXTOVAR.Visible = False
  grid_fac.Col = grid_fac.Col + 2
  grid_fac.SetFocus
  Exit Sub
End If
If grid_fac.Col = 2 Then
  TEXTOVAR.Visible = False
  grid_fac.SetFocus
'  FrmPlan.cmdnuevoes.SetFocus
  Exit Sub
End If


TEXTOVAR.Visible = False
grid_fac.SetFocus
Exit Sub
    SQ_OPER = 1
    PUB_CUENTA = TEXTOVAR.Text
    PUB_CODCIA = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
         MsgBox "CUENTA NO EXISTE ...", 48, Pub_Titulo
         Azul3 TEXTOVAR, TEXTOVAR
         GoTo fin
    Else
     If Val(com_llave!com_nivel) <> Val(cop_llave!cop_nivel_max) Then
         MsgBox "Cuenta es validad, pero no es Analitica ...", 48, Pub_Titulo
         Azul3 TEXTOVAR, TEXTOVAR
         GoTo fin
     End If
     grid_fac.TextMatrix(grid_fac.Row, 2) = Trim(com_llave!com_DESCRIPCION)
     grid_fac.TextMatrix(grid_fac.Row, 1) = Trim(com_llave!com_cuenta)
    End If
    
 
 

fin:

End Sub

Private Sub textovar_LostFocus()
'TEXTOVAR.Visible = False
If TEXTOVAR.Visible Then
 '  TEXTOVAR.Visible = False
    '   grid_fac.Row = wfila_act
'   grid_fac.SetFocus
   Exit Sub
End If

End Sub
Public Sub BUSCAR_CTA(WTIPO As Integer)
Dim WCUENTA As TextBox
Dim wgrupo As String
Dim wq_cuenta As String

LK_TABLA = "BUSCAR2"
If WTIPO = 1 Then
 If TEXTOVAR.Text = "*" Then
  wgrupo = "" 'Trim(i_cuenta.text)
  archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_CUENTA >= '" & wgrupo & "'  ORDER BY COM_CUENTA"
 Else
 TEXTOVAR.Text = Mid(TEXTOVAR.Text, 2, Len(TEXTOVAR.Text))
 wgrupo = Trim(TEXTOVAR.Text)
 If Val(wgrupo) = 0 Then Exit Sub
 archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_CUENTA >= '" & wgrupo & "' AND COM_CUENTA < '" & Trim(Str(Val(wgrupo) + 1)) & "'  ORDER BY COM_CUENTA"
 End If
Else
 If TEXTOVAR.Text = "*" Then
  wgrupo = "" 'Trim(i_cuenta.text)
  archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_CUENTA >= '" & wgrupo & "'  ORDER BY COM_CUENTA"
 Else
 TEXTOVAR.Text = Mid(TEXTOVAR.Text, 2, Len(TEXTOVAR.Text))
 wgrupo = Trim(TEXTOVAR.Text)
 If Val(wgrupo) = 0 Then Exit Sub
 archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_CUENTA >= '" & wgrupo & "' AND COM_CUENTA < '" & Trim(Str(Val(wgrupo) + 1)) & "'  ORDER BY COM_CUENTA"
 End If
End If
Load frmBuscacta
frmBuscacta.lbltabla.Caption = LK_TABLA
frmBuscacta.Show 1
wq_cuenta = Trim(frmBuscacta.tcuenta)
If wq_cuenta <> "" Then
  TEXTOVAR.Text = Trim(frmBuscacta.tcuenta)
End If
Unload frmBuscacta
If wq_cuenta <> "" Then
   If WTIPO = 1 Then
 '    i_cuenta_KeyPress 13
   Else
     textovar_KeyPress 13
   End If
Else
  Azul3 TEXTOVAR, TEXTOVAR
End If


End Sub
Private Sub tglosa_Change()
If Trim(listv.Text) = "" Then
  tglosa.Text = ""
Else
 cmdgrabar.Enabled = True
End If
End Sub

Private Sub tglosa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If grid_fac.Rows <= 1 Then
    cmdnuevoes.SetFocus
    Exit Sub
  End If
  
  
    grid_fac.Col = 0
    grid_fac.Row = 1
    grid_fac.SetFocus
End If
End Sub

Public Sub cabegrid()
grid_fac.Clear
grid_fac.Cols = 5
grid_fac.Rows = 2
grid_fac.TextMatrix(0, 0) = "Cod.Cta."
grid_fac.TextMatrix(0, 1) = "Descripción de Cta."
grid_fac.TextMatrix(0, 2) = "D/H"
grid_fac.TextMatrix(0, 4) = "Busqu."
grid_fac.ColWidth(0) = 800 ' cuenta
grid_fac.ColWidth(1) = 2400 ' descrip - cuenta
grid_fac.ColWidth(2) = 600 ' d/h
grid_fac.ColWidth(3) = 0 ' secuencia
grid_fac.ColWidth(4) = 600 ' FLAG BUSQUEDA

End Sub
Public Sub cabegriddes()
grid_des.Clear
grid_des.Cols = 5
grid_des.Rows = 2
grid_des.TextMatrix(0, 0) = "Item"
grid_des.TextMatrix(0, 1) = "Cuenta"
grid_des.TextMatrix(0, 2) = "Descripción"
grid_des.TextMatrix(0, 3) = "D/H"
grid_des.TextMatrix(0, 4) = "(%)"
grid_des.ColWidth(0) = 400 ' cuenta
grid_des.ColWidth(1) = 900 ' descrip - cuenta
grid_des.ColWidth(2) = 2300 ' d/h
grid_des.ColWidth(3) = 400 ' secuencia
grid_des.ColWidth(4) = 600 ' FLAG BUSQUEDA

End Sub

Public Sub llena_data(wnume As Integer)
loc_key = 9
cabegrid
cabegriddes
PSPL_LLAVE(0) = LK_CODCIA
PSPL_LLAVE(1) = wnume
plt_llave.Requery
If plt_llave.EOF Then Exit Sub
fila = 0
plt_llave.MoveNext
tglosa.Text = ""
If Not plt_llave.EOF Then tglosa.Text = Trim(Nulo_Valors(plt_llave!PLT_GLOSA))
fila = 0
grid_fac.Rows = 1
Do Until plt_llave.EOF
    fila = fila + 1
    grid_fac.Rows = grid_fac.Rows + 1
    grid_fac.Row = fila
    grid_fac.TextMatrix(fila, 0) = Trim(plt_llave!PLT_CUENTA)
    SQ_OPER = 1
    PUB_CUENTA = Trim(plt_llave!PLT_CUENTA)
    PUB_CODCIA = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
       grid_fac.TextMatrix(fila, 1) = "Verificar esta Cuenta..."
    Else
       grid_fac.TextMatrix(fila, 1) = com_llave!com_DESCRIPCION
    End If
    
    grid_fac.Col = 2
    grid_fac.CellAlignment = 4
    grid_fac.TextMatrix(fila, 2) = Trim(plt_llave!PLT_DH)
    grid_fac.TextMatrix(fila, 3) = plt_llave!PLT_SECUENCIA
    grid_fac.Col = 4
    grid_fac.CellAlignment = 4
    If Trim(plt_llave!PLT_NOMBRE) <> "" Then
      grid_fac.TextMatrix(fila, 4) = Trim(plt_llave!PLT_NOMBRE)
    Else
      grid_fac.TextMatrix(fila, 4) = ""
    End If
    plt_llave.MoveNext
Loop
lvoucher.Caption = Trim(Left(listv.Text, 40))
lcodigo.Caption = listv.ListIndex
cmdgrabar.Enabled = False
loc_key = 0

End Sub


Public Sub llena_destinos(WCUENTA As String)
If loc_key = 9 Then Exit Sub
If Trim(WCUENTA) = "" Then Exit Sub
cabegriddes
SQ_OPER = 1
PUB_CUENTA = WCUENTA
PUB_CODCIA = LK_CODCIA
LEER_COM_LLAVE
If com_llave.EOF Then
  MsgBox "Verificar Cuenta en Plan de cuentas..!", 48, Pub_Titulo
  Exit Sub
End If
If Val(com_llave!com_nivel) <> Val(cop_llave!cop_nivel_max) Then
  Exit Sub
End If
grid_des.Rows = 1
If Trim(com_llave!com_cuenta_AUTO_H) <> "" Then
  grid_des.Rows = grid_des.Rows + 1
  grid_des.TextMatrix(grid_des.Rows - 1, 0) = grid_des.Rows - 2
  grid_des.TextMatrix(grid_des.Rows - 1, 1) = com_llave!com_cuenta_AUTO_H
  grid_des.TextMatrix(grid_des.Rows - 1, 2) = ""
  grid_des.TextMatrix(grid_des.Rows - 1, 3) = "H"
  grid_des.TextMatrix(grid_des.Rows - 1, 4) = "100.00"
End If

If Trim(com_llave!com_cuenta_AUTOM_D) <> "" Then
  grid_des.Rows = grid_des.Rows + 1
  grid_des.TextMatrix(grid_des.Rows - 1, 0) = grid_des.Rows - 2
  grid_des.TextMatrix(grid_des.Rows - 1, 1) = com_llave!com_cuenta_AUTOM_D
  grid_des.TextMatrix(grid_des.Rows - 1, 2) = ""
  grid_des.TextMatrix(grid_des.Rows - 1, 3) = "D"
  grid_des.TextMatrix(grid_des.Rows - 1, 4) = Format(com_llave!COM_POR_AUTOM_D, "0.00")
End If
If Trim(com_llave!COM_CUENTA_AUTOM_D2) <> "" Then
  grid_des.Rows = grid_des.Rows + 1
  grid_des.TextMatrix(grid_des.Rows - 1, 0) = grid_des.Rows - 2
  grid_des.TextMatrix(grid_des.Rows - 1, 1) = com_llave!COM_CUENTA_AUTOM_D2
  grid_des.TextMatrix(grid_des.Rows - 1, 2) = ""
  grid_des.TextMatrix(grid_des.Rows - 1, 3) = "D"
  grid_des.TextMatrix(grid_des.Rows - 1, 4) = Format(com_llave!COM_POR_AUTOM_D2, "0.00")
End If
If Trim(com_llave!COM_CUENTA_AUTOM_D3) <> "" Then
  grid_des.Rows = grid_des.Rows + 1
  grid_des.TextMatrix(grid_des.Rows - 1, 0) = grid_des.Rows - 2
  grid_des.TextMatrix(grid_des.Rows - 1, 1) = com_llave!COM_CUENTA_AUTOM_D3
  grid_des.TextMatrix(grid_des.Rows - 1, 2) = ""
  grid_des.TextMatrix(grid_des.Rows - 1, 3) = "D"
  grid_des.TextMatrix(grid_des.Rows - 1, 4) = Format(com_llave!COM_POR_AUTOM_D3, "0.00")
End If
If Trim(com_llave!COM_CUENTA_AUTOM_D4) <> "" Then
  grid_des.Rows = grid_des.Rows + 1
  grid_des.TextMatrix(grid_des.Rows - 1, 0) = grid_des.Rows - 2
  grid_des.TextMatrix(grid_des.Rows - 1, 1) = com_llave!COM_CUENTA_AUTOM_D4
  grid_des.TextMatrix(grid_des.Rows - 1, 2) = ""
  grid_des.TextMatrix(grid_des.Rows - 1, 3) = "D"
  grid_des.TextMatrix(grid_des.Rows - 1, 4) = Format(com_llave!COM_POR_AUTOM_D4, "0.00")
End If
If Trim(com_llave!COM_CUENTA_AUTOM_D5) <> "" Then
  grid_des.Rows = grid_des.Rows + 1
  grid_des.TextMatrix(grid_des.Rows - 1, 0) = grid_des.Rows - 2
  grid_des.TextMatrix(grid_des.Rows - 1, 1) = com_llave!COM_CUENTA_AUTOM_D5
  grid_des.TextMatrix(grid_des.Rows - 1, 2) = ""
  grid_des.TextMatrix(grid_des.Rows - 1, 3) = "D"
  grid_des.TextMatrix(grid_des.Rows - 1, 4) = Format(com_llave!COM_POR_AUTOM_D5, "0.00")
End If



For fila = 1 To grid_des.Rows - 1
    SQ_OPER = 1
    PUB_CUENTA = grid_des.TextMatrix(fila, 1)
    PUB_CODCIA = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
      MsgBox "Verificar Cuenta de destino en Plan de cuentas..!" + PUB_CUENTA, 48, Pub_Titulo
    End If
    grid_des.TextMatrix(fila, 2) = com_llave!com_DESCRIPCION
Next fila

End Sub
