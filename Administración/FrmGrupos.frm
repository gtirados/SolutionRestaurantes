VERSION 5.00
Begin VB.Form FrmGrupos 
   Caption         =   "Defenir Grupo de Trabajos x Transacción"
   ClientHeight    =   4980
   ClientLeft      =   1050
   ClientTop       =   1155
   ClientWidth     =   8880
   Icon            =   "FrmGrupos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   8880
   Begin VB.Frame fra1 
      Caption         =   "Transacción "
      Height          =   4455
      Left            =   3840
      TabIndex        =   1
      Top             =   0
      Width           =   4695
      Begin VB.Frame fraUsu 
         Caption         =   "Usuarios Disponibles"
         Height          =   2775
         Left            =   1200
         TabIndex        =   14
         Top             =   1680
         Visible         =   0   'False
         Width           =   3495
         Begin VB.ListBox listUsu 
            Height          =   2205
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label3 
            Caption         =   "[Enter]= Agregar"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   2445
            Width           =   1695
         End
      End
      Begin VB.Frame fraGrupo 
         Caption         =   "Grupos Disponibles"
         Height          =   2775
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   2895
         Begin VB.ListBox listgrupos 
            Height          =   2205
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "[Enter]= Agrega          [Insert]= Nuevo"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   2450
            Width           =   2655
         End
      End
      Begin VB.ListBox GRU_USU 
         BackColor       =   &H00C0C0C0&
         Height          =   1230
         Left            =   120
         TabIndex        =   7
         Top             =   3000
         Width           =   3135
      End
      Begin VB.ListBox lisdisponible 
         Height          =   1425
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "[DEL] = Quitar Usuario"
         Height          =   495
         Index           =   5
         Left            =   3360
         TabIndex        =   17
         Top             =   3720
         Width           =   1095
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "Grupos de trabajo con Aceeso"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "[Insert] = Agregar                     [Del] = Quitar"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Usuarios con Aceesos"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label lbltra 
         Alignment       =   2  'Center
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
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "[F4] = Agregar Usuario a Grupo"
         Height          =   495
         Index           =   4
         Left            =   3360
         TabIndex        =   13
         Top             =   840
         Width           =   1215
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "Ce&rrar"
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   4560
      Width           =   1335
   End
   Begin VB.ListBox listtra 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Lista de Transacciones "
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
      Index           =   3
      Left            =   0
      TabIndex        =   11
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "FrmGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset
Dim PS_REP02 As rdoQuery
Dim llave_rep02 As rdoResultset


Private Sub cmdcerrar_Click()
Unload FrmGrupos
End Sub

Private Sub Form_Load()
CenterMe FrmGrupos
pub_cadena = "SELECT TRA_KEY,TRA_DESCRIPCION,TRA_GRU1,TRA_GRU2,TRA_GRU3,TRA_GRU4,TRA_GRU5,TRA_GRU6,TRA_GRU7,TRA_GRU8,TRA_GRU9,TRA_GRU10 FROM TRANSACCION WHERE TRA_KEY = ? ORDER BY TRA_KEY"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM USUARIOS WHERE USU_KEY = ? ORDER BY USU_KEY"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurValues)

LLENA_USU
LLENA_TRA
PUB_CODCIA = "00"
LLENA_GRUPOS listgrupos, 80

End Sub

Public Sub LLENA_TRA()
Dim wtra As String * 5
Dim wnombre As String * 25
wtra = String(5, " ")
wnombre = String(25, " ")
lis_tra.Requery
listtra.Clear
Do Until lis_tra.EOF
    If gen!gen_bloqueo = "A" And lis_tra!TRA_KEY = 1409 Then GoTo SAL
    wtra = Trim(lis_tra!TRA_KEY)
    wnombre = Trim(lis_tra!tra_descripcion)
    listtra.AddItem wnombre & " " & wtra
SAL:
    lis_tra.MoveNext
Loop
End Sub
Public Sub LLENA_USU()
Dim wtra As String * 5
Dim wnombre As String * 30
'wtra = String(5, " ")
wnombre = String(30, " ")
usu.Requery
listUsu.Clear
Do Until usu.EOF
    'wtra = Trim(usu!USU_NOMBRE)
    If Trim(usu!usu_key) = "ADMIN" And LK_CODUSU <> "ADMIN" Then
    Else
     wnombre = Trim(usu!USU_NOMBRE)
     listUsu.AddItem wnombre & String(60, " ") & usu!usu_key
    End If
    usu.MoveNext
Loop
End Sub

Public Sub LLENA_GRUPOS(cont As ListBox, tip As Integer)
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.Clear
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!TAB_NOMLARGO & String(50, "  ") & tab_mayor!TAB_NUMTAB
        tab_mayor.MoveNext
    Loop
End Sub

Private Sub GRU_USU_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
  If Trim(GRU_USU.text) = "" Then
    Exit Sub
  End If
  pub_mensaje = "Quitar el Usuario del Grupo  ¿Desea Continuar... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then
   Exit Sub
  End If
  PS_REP02(0) = Trim(Right(GRU_USU.text, 10))
  llave_rep02.Requery
  If llave_rep02.EOF Then
    MsgBox "Intente Nuevamente.", 48, Pub_Titulo
    Exit Sub
  End If
  Wflag = 0
  llave_rep02.Edit
  For fila = 4 To 13
     If Val(llave_rep02.rdoColumns(fila)) = Val(Trim(Right(lisdisponible.text, 5))) Then
        Wflag = 1
        llave_rep02.rdoColumns(fila) = 0
        End If
  Next fila
  If Wflag = 0 Then
      llave_rep02.CancelUpdate
      MsgBox "Intente Nuevamente", 48, Pub_Titulo
  End If
  GRU_USU.RemoveItem GRU_USU.ListIndex
  llave_rep02.Update
End If

End Sub

Private Sub GRU_USU_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then FrmGrupos.lisdisponible.SetFocus
End Sub

Private Sub lisdisponible_Click()
If Val(Trim(Right(lisdisponible.text, 4))) = 0 Then Exit Sub
TOMA_USUARIOS Val(Trim(Right(lisdisponible.text, 4)))
End Sub

Private Sub lisdisponible_GotFocus()
fraGrupo.Visible = False
End Sub

Private Sub lisdisponible_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then
  fraUsu.Visible = True
  listUsu.SetFocus
  Exit Sub
End If

If KeyCode = 45 Then
  fraGrupo.Visible = True
  listgrupos.SetFocus
End If
If KeyCode = 46 Then
  If Trim(lisdisponible.text) = "" Then
   Exit Sub
  End If
  pub_mensaje = "Quitar el Grupo de Trabajo  ¿Desea Continuar... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then
   Exit Sub
  End If
  Wflag = 0
  For fila = 2 To 11
     If Val(llave_rep01.rdoColumns(fila)) = Val(Trim(Right(lisdisponible.text, 5))) Then
        Wflag = 1
        llave_rep01.Edit
        llave_rep01.rdoColumns(fila) = 0
        llave_rep01.Update
        lisdisponible.RemoveItem lisdisponible.ListIndex
        Exit For
      End If
  Next fila
  If Wflag = 0 Then
      MsgBox "Intente Nuevamente", 48, Pub_Titulo
  End If

End If


End Sub

Private Sub lisdisponible_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then listtra.SetFocus
If KeyAscii = 13 Then
  GRU_USU.SetFocus
  If GRU_USU.ListCount > 0 Then
   GRU_USU.ListIndex = 0
  End If
End If
End Sub

Private Sub listgrupos_DblClick()
listgrupos_KeyPress 13
End Sub

Private Sub listgrupos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 Then
Dim wpos As Integer
'If KeyCode <> 45 Then
'  Exit Sub
'End If
wpos = listgrupos.ListIndex
PUB_TIPREG = 80
PUB_CODCIA = "00"
Load FrmDatArti
FrmDatArti.Caption = "Grupos de Trabajo"
FrmDatArti.Show 1
DoEvents
LLENA_GRUPOS listgrupos, 80
listgrupos.SetFocus
End If
End Sub

Private Sub listgrupos_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  fraGrupo.Visible = False
  lisdisponible.SetFocus
End If
If KeyAscii = 13 Then
Dim Wflag As Integer
For fila = 0 To lisdisponible.ListCount - 1
lisdisponible.ListIndex = fila
If Trim(Right(lisdisponible.text, 5)) = Trim(Right(listgrupos.text, 5)) Then
  MsgBox " Grupo esta en la Relación.", 48, Pub_Titulo
  Exit Sub
End If
Next fila
Wflag = 0
For fila = 2 To 11
 If Val(llave_rep01.rdoColumns(fila)) = 0 Then
    Wflag = 1
    llave_rep01.Edit
    llave_rep01.rdoColumns(fila) = Val(Trim(Right(listgrupos.text, 5)))
    llave_rep01.Update
    lisdisponible.AddItem listgrupos.text
    Exit For
  End If
Next fila
If Wflag = 0 Then
  MsgBox "No Procede,  tope Maximo de Grupo", 48, Pub_Titulo
End If
fraGrupo.Visible = False
lisdisponible.SetFocus
End If

End Sub

Private Sub listtra_Click()
Dim wnombre As String
GRU_USU.Clear
PS_REP01(0) = Val(Trim(Right(listtra.text, 5)))
llave_rep01.Requery
If llave_rep01.EOF Then
  MsgBox "Intente Nuevamente.", 48, Pub_Titulo
  Exit Sub
End If
lbltra.Caption = llave_rep01!tra_descripcion
lisdisponible.Clear

If Val(llave_rep01!TRA_GRU1) <> 0 Then lisdisponible.AddItem BUSCA_GRUPO(Val(llave_rep01!TRA_GRU1)) & String(60, " ") & Str(llave_rep01!TRA_GRU1)
If Val(llave_rep01!TRA_GRU2) <> 0 Then lisdisponible.AddItem BUSCA_GRUPO(Val(llave_rep01!TRA_GRU2)) & String(60, " ") & Str(llave_rep01!TRA_GRU2)
If Val(llave_rep01!TRA_GRU3) <> 0 Then lisdisponible.AddItem BUSCA_GRUPO(Val(llave_rep01!TRA_GRU3)) & String(60, " ") & Str(llave_rep01!TRA_GRU3)
If Val(llave_rep01!TRA_GRU4) <> 0 Then lisdisponible.AddItem BUSCA_GRUPO(Val(llave_rep01!TRA_GRU4)) & String(60, " ") & Str(llave_rep01!TRA_GRU4)
If Val(llave_rep01!TRA_GRU5) <> 0 Then lisdisponible.AddItem BUSCA_GRUPO(Val(llave_rep01!TRA_GRU5)) & String(60, " ") & Str(llave_rep01!TRA_GRU5)
If Val(llave_rep01!TRA_GRU6) <> 0 Then lisdisponible.AddItem BUSCA_GRUPO(Val(llave_rep01!TRA_GRU6)) & String(60, " ") & Str(llave_rep01!TRA_GRU6)
If Val(llave_rep01!TRA_GRU7) <> 0 Then lisdisponible.AddItem BUSCA_GRUPO(Val(llave_rep01!TRA_GRU7)) & String(60, " ") & Str(llave_rep01!TRA_GRU7)
If Val(llave_rep01!TRA_GRU8) <> 0 Then lisdisponible.AddItem BUSCA_GRUPO(Val(llave_rep01!TRA_GRU8)) & String(60, " ") & Str(llave_rep01!TRA_GRU8)
If Val(llave_rep01!TRA_GRU8) <> 0 Then lisdisponible.AddItem BUSCA_GRUPO(Val(llave_rep01!TRA_GRU9)) & String(60, " ") & Str(llave_rep01!TRA_GRU9)
If Val(llave_rep01!TRA_GRU10) <> 0 Then lisdisponible.AddItem BUSCA_GRUPO(Val(llave_rep01!TRA_GRU10)) & String(60, " ") & Str(llave_rep01!TRA_GRU10)


End Sub

Public Function BUSCA_GRUPO(WCODI As Integer) As String
PUB_TIPREG = 80
PUB_NUMTAB = WCODI
PUB_CODCIA = "00"
SQ_OPER = 1
LEER_TAB_LLAVE
If tab_llave.EOF Then
 BUSCA_GRUPO = "* Disponible *"
 Exit Function
End If
BUSCA_GRUPO = Trim(tab_llave!TAB_NOMLARGO)
End Function

Public Function TOMA_USUARIOS(WCOD As Integer)
usu.Requery
GRU_USU.Clear
Do Until usu.EOF
If Trim(usu!usu_key) = "ADMIN" And LK_CODUSU <> "ADMIN" Then GoTo OTRO
If usu!USU_GRUPO1 = WCOD Then GRU_USU.AddItem usu!USU_NOMBRE & String(60, " ") & usu!usu_key: GoTo OTRO
If usu!USU_GRUPO2 = WCOD Then GRU_USU.AddItem usu!USU_NOMBRE & String(60, " ") & usu!usu_key: GoTo OTRO
If usu!USU_GRUPO3 = WCOD Then GRU_USU.AddItem usu!USU_NOMBRE & String(60, " ") & usu!usu_key: GoTo OTRO
If usu!USU_GRUPO4 = WCOD Then GRU_USU.AddItem usu!USU_NOMBRE & String(60, " ") & usu!usu_key: GoTo OTRO
If usu!USU_GRUPO5 = WCOD Then GRU_USU.AddItem usu!USU_NOMBRE & String(60, " ") & usu!usu_key: GoTo OTRO
If usu!USU_GRUPO6 = WCOD Then GRU_USU.AddItem usu!USU_NOMBRE & String(60, " ") & usu!usu_key: GoTo OTRO
If usu!USU_GRUPO7 = WCOD Then GRU_USU.AddItem usu!USU_NOMBRE & String(60, " ") & usu!usu_key: GoTo OTRO
If usu!USU_GRUPO8 = WCOD Then GRU_USU.AddItem usu!USU_NOMBRE & String(60, " ") & usu!usu_key: GoTo OTRO
If usu!USU_GRUPO9 = WCOD Then GRU_USU.AddItem usu!USU_NOMBRE & String(60, " ") & usu!usu_key: GoTo OTRO
If usu!USU_GRUPO10 = WCOD Then GRU_USU.AddItem usu!USU_NOMBRE & String(60, " ") & usu!usu_key: GoTo OTRO
OTRO:
usu.MoveNext
Loop
End Function

Private Sub listtra_GotFocus()
fraGrupo.Visible = False
End Sub

Private Sub listtra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 lisdisponible.SetFocus
 If lisdisponible.ListCount > 0 Then
   lisdisponible.ListIndex = 0
 End If
End If
End Sub

Private Sub listUsu_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  fraUsu.Visible = False
  lisdisponible.SetFocus
End If
If KeyAscii = 13 Then
Dim Wflag As Integer
For fila = 0 To GRU_USU.ListCount - 1
GRU_USU.ListIndex = fila
If Trim(Right(listUsu.text, 10)) = Trim(Right(GRU_USU.text, 10)) Then
  MsgBox " Usuario esta en la Relación.", 48, Pub_Titulo
  Exit Sub
End If
Next fila
PS_REP02(0) = Trim(Right(listUsu.text, 10))
llave_rep02.Requery
If llave_rep02.EOF Then
  MsgBox "Intente Nuevamente.", 48, Pub_Titulo
  Exit Sub
End If

Wflag = 0
For fila = 4 To 13
 If Val(llave_rep02.rdoColumns(fila)) = 0 Then
    Wflag = 1
    llave_rep02.Edit
    llave_rep02.rdoColumns(fila) = Val(Trim(Right(lisdisponible.text, 4)))
    llave_rep02.Update
    Exit For
  End If
Next fila
If Wflag = 0 Then
  MsgBox "No Procede,  tope Maximo de Grupo", 48, Pub_Titulo
End If
If Val(Trim(Right(lisdisponible.text, 4))) = 0 Then Exit Sub
 TOMA_USUARIOS Val(Trim(Right(lisdisponible.text, 4)))
End If


If KeyAscii = 13 Then
  fraUsu.Visible = False
 lisdisponible.SetFocus
End If

End Sub
