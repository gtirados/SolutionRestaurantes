VERSION 5.00
Begin VB.Form frmTablas 
   Caption         =   "Mantenimiento de Tablas"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9045
   ControlBox      =   0   'False
   Icon            =   "frmTablas.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   9045
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "Ce&rrar"
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones:"
      Height          =   3255
      Left            =   3960
      TabIndex        =   5
      Top             =   480
      Width           =   3135
      Begin VB.ListBox fami 
         Height          =   2595
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "No hay Opciones."
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lblfami 
         Caption         =   "Seleccione una Familia :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdmante 
      Caption         =   "Dar &Mantenimiento"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Frame fratablas 
      Caption         =   "Seleccionar una Tabla :"
      Height          =   4095
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   3255
      Begin VB.ListBox lista 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3660
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmTablas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Wflag   As String * 1
Dim LOC_TIPREG As Integer
Dim LOC_CODCIA As String * 2

Private Sub cmdcerrar_Click()
Unload frmTablas
End Sub

Private Sub cmdmante_Click()
Dim wfami
If LOC_TIPREG = -10 Then
PUB_TIPREG = LOC_TIPREG
PUB_CODCIA = LOC_CODCIA
Load FrmDatArti
FrmDatArti.Caption = "Mantenimiento de Transportistas"
FrmDatArti.Show 1
lista.SetFocus
DoEvents

Exit Sub
End If
If LOC_TIPREG = 123 Then
  If Trim(fami.text) = "" Then
    MsgBox "Definir una Familia para relacionar.", 48, Pub_Titulo
    fami.SetFocus
    Exit Sub
  End If
  wfami = Val(Trim(Right(fami.text, 6)))
  PUB_CODART = wfami
End If
PUB_TIPREG = LOC_TIPREG
PUB_CODCIA = LOC_CODCIA
Load FrmDatArti
FrmDatArti.Caption = Trim(Left(lista.text, 30)) & " TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
lista.SetFocus
DoEvents
End Sub

Private Sub fami_DblClick()
cmdmante_Click
End Sub

Private Sub fami_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  cmdmante_Click
End If
If KeyAscii = 27 Then
  lista.SetFocus
End If

End Sub

Private Sub Form_Load()
Wflag = "A"
lista.AddItem BUSCA_ETIQUETA(10) + String(80, " ") + "20"
lista.AddItem BUSCA_ETIQUETA(11) + String(80, " ") + "30"
lista.AddItem BUSCA_ETIQUETA(12) + String(80, " ") + "35"
lista.AddItem "Familias " + String(80, " ") + "122"
lista.AddItem "Sub-Familias" + String(80, " ") + "123"
lista.AddItem "Grupos de Trabajos" + String(80, " ") + "80"
lista.AddItem "Calidades" + String(80, " ") + "2"
lista.AddItem "Tipo de Doc. SUNAT " + String(80, " ") + "50"
lista.AddItem "Tipos de Bloques " + String(80, " ") + "199"
lista.AddItem "Tipos de Cta. Contables " + String(80, " ") + "16"
lista.AddItem "Datos de Transportista " + String(80, " ") + "-10"

SQ_OPER = 1
PUB_TIPREG = 300
PUB_CODCIA = LK_CODCIA
PUB_NUMTAB = 13
LEER_TAB_LLAVE
If tab_llave.EOF Then
Else
   lista.AddItem Trim(tab_llave!TAB_NOMLARGO) + String(80, " ") + "222"
End If

PUB_NUMTAB = 14
LEER_TAB_LLAVE
If tab_llave.EOF Then
Else
   lista.AddItem Trim(tab_llave!TAB_NOMLARGO) + String(80, " ") + "333"
End If
Wflag = ""
End Sub

Private Sub lista_Click()
If Wflag = "A" Then
 Exit Sub
End If
LOC_TIPREG = Val(Right(lista.text, 6))
If LOC_TIPREG = 20 Or LOC_TIPREG = 30 Or LOC_TIPREG = 35 Or LOC_TIPREG = 80 Or LOC_TIPREG = 50 Then
 LOC_CODCIA = "00"
Else
 LOC_CODCIA = LK_CODCIA
End If
lblfami.Visible = False
fami.Visible = False
If LOC_TIPREG = 123 Then
  lblfami.Visible = True
  PUB_CODCIA = LK_CODCIA
  fami.Visible = True
  LLENADOS fami, 122
End If

End Sub

Private Sub lista_DblClick()
cmdmante_Click
End Sub
Public Sub LLENADOS(cont As ListBox, tip As Integer)
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
       cont.AddItem tab_mayor!TAB_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
       tab_mayor.MoveNext
    Loop
End Sub

Private Sub lista_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  cmdmante_Click
End If

End Sub
