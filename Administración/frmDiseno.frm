VERSION 5.00
Begin VB.Form frmDiseno 
   Caption         =   "Diseño de Transacciones"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   1095
   ClientWidth     =   9480
   Icon            =   "frmDiseno.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5805
   ScaleWidth      =   9480
   Begin VB.Frame fop3 
      Caption         =   "Opcionales"
      Height          =   3495
      Left            =   7200
      TabIndex        =   53
      Top             =   480
      Width           =   2175
      Begin VB.TextBox txtplaca 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   61
         Top             =   2730
         Width           =   1815
      End
      Begin VB.TextBox txtruc 
         Height          =   285
         Left            =   120
         MaxLength       =   8
         TabIndex        =   59
         Top             =   2130
         Width           =   1815
      End
      Begin VB.TextBox txtdirec 
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   57
         Top             =   1410
         Width           =   1815
      End
      Begin VB.TextBox txtnombre 
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   55
         Top             =   810
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Placa"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   62
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "R.U.C."
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   60
         Top             =   1920
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Dirección"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   58
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   56
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "Datos del Transportista"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   54
         Top             =   360
         Width           =   1845
      End
   End
   Begin VB.ListBox listatra 
      BackColor       =   &H00C0C0C0&
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
      Height          =   2595
      Left            =   1920
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame fop2 
      Caption         =   "Opciones para Reporte :"
      Height          =   3495
      Left            =   0
      TabIndex        =   18
      Top             =   480
      Width           =   7095
      Begin VB.TextBox OBJETO 
         Appearance      =   0  'Flat
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   50
         Text            =   "OBJETO"
         Top             =   600
         Visible         =   0   'False
         Width           =   2600
      End
      Begin VB.TextBox OBJETO 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   49
         Text            =   "OBJETO"
         Top             =   960
         Visible         =   0   'False
         Width           =   2600
      End
      Begin VB.TextBox OBJETO 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   48
         Text            =   "OBJETO"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2600
      End
      Begin VB.TextBox OBJETO 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   47
         Text            =   "OBJETO"
         Top             =   1680
         Visible         =   0   'False
         Width           =   2600
      End
      Begin VB.TextBox OBJETO 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   46
         Text            =   "OBJETO"
         Top             =   2040
         Visible         =   0   'False
         Width           =   2600
      End
      Begin VB.TextBox OBJETO 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "OBJETO"
         Top             =   2400
         Visible         =   0   'False
         Width           =   2600
      End
      Begin VB.TextBox OBJETO 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "OBJETO"
         Top             =   2760
         Visible         =   0   'False
         Width           =   2600
      End
      Begin VB.TextBox OBJETO 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "OBJETO"
         Top             =   3120
         Visible         =   0   'False
         Width           =   2600
      End
      Begin VB.TextBox OBJETO 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   42
         Text            =   "OBJETO"
         Top             =   600
         Visible         =   0   'False
         Width           =   2600
      End
      Begin VB.TextBox OBJETO 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "OBJETO"
         Top             =   960
         Visible         =   0   'False
         Width           =   2600
      End
      Begin VB.TextBox OBJETO 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "OBJETO"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2600
      End
      Begin VB.TextBox OBJETO 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "OBJETO"
         Top             =   1680
         Visible         =   0   'False
         Width           =   2600
      End
      Begin VB.TextBox OBJETO 
         Appearance      =   0  'Flat
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "OBJETO"
         Top             =   2040
         Visible         =   0   'False
         Width           =   2600
      End
      Begin VB.TextBox OBJETO 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "OBJETO"
         Top             =   2400
         Visible         =   0   'False
         Width           =   2600
      End
      Begin VB.TextBox OBJETO 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "OBJETO"
         Top             =   2760
         Visible         =   0   'False
         Width           =   2600
      End
      Begin VB.TextBox OBJETO 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "OBJETO"
         Top             =   3120
         Visible         =   0   'False
         Width           =   2600
      End
      Begin VB.CheckBox cheactivo 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox cheactivo 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox cheactivo 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox cheactivo 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox cheactivo 
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   30
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox cheactivo 
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   29
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox cheactivo 
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   28
         Top             =   2760
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox cheactivo 
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   27
         Top             =   3120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox cheactivo 
         Height          =   255
         Index           =   8
         Left            =   3960
         TabIndex        =   26
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox cheactivo 
         Height          =   255
         Index           =   9
         Left            =   3960
         TabIndex        =   25
         Top             =   960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox cheactivo 
         Height          =   255
         Index           =   10
         Left            =   3960
         TabIndex        =   24
         Top             =   1320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox cheactivo 
         Height          =   255
         Index           =   11
         Left            =   3960
         TabIndex        =   23
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox cheactivo 
         Height          =   255
         Index           =   12
         Left            =   3960
         TabIndex        =   22
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox cheactivo 
         Height          =   255
         Index           =   13
         Left            =   3960
         TabIndex        =   21
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox cheactivo 
         Height          =   255
         Index           =   14
         Left            =   3960
         TabIndex        =   20
         Top             =   2760
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox cheactivo 
         Height          =   255
         Index           =   15
         Left            =   3960
         TabIndex        =   19
         Top             =   3120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Activo    Codigo          Descripcion de Campo"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   3405
      End
      Begin VB.Label Label1 
         Caption         =   "Activo    Codigo          Descripcion de Campo"
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   51
         Top             =   240
         Width           =   3285
      End
   End
   Begin VB.CheckBox cherepactivo 
      Caption         =   "Reporte Activado"
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
      Left            =   6960
      TabIndex        =   17
      Top             =   120
      Width           =   2295
   End
   Begin VB.Frame fop1 
      Caption         =   "Opciones para Impresión"
      Height          =   1695
      Left            =   0
      TabIndex        =   6
      Top             =   3960
      Width           =   7455
      Begin VB.CommandButton cmdGene 
         Caption         =   "&Generar Hoja de Calculo"
         Height          =   350
         Left            =   2880
         TabIndex        =   16
         Top             =   1200
         Width           =   2895
      End
      Begin VB.CommandButton cmdExaminar 
         Caption         =   "&Examinar"
         Height          =   350
         Left            =   6240
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox cheimp 
         Caption         =   "Directo a Impresora"
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtruta 
         Height          =   285
         Left            =   2640
         TabIndex        =   12
         Top             =   840
         Width           =   3495
      End
      Begin VB.CheckBox chetra 
         Caption         =   "Cod. Transac.  ( WF603 )"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CheckBox cheusu 
         Caption         =   "Usuario  ( WF602 )"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chedia 
         Caption         =   "Fecha del Sistema  ( WF601 )"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2415
      End
      Begin VB.CheckBox checia 
         Caption         =   "Compañia ( WF600 )"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Ruta de Archivo  Hoja de Calculo Excel."
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   11
         Top             =   600
         Width           =   3285
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   7560
      TabIndex        =   3
      Top             =   3960
      Width           =   1815
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancelar  / Cerrar"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton Grabar 
         Caption         =   "&Grabar "
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox trans 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   600
      X2              =   7440
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Label transa 
      AutoSize        =   -1  'True
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
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Transacion :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   885
   End
End
Attribute VB_Name = "frmDiseno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim C_Objeto As Integer
Dim ACTIVO As Integer
Dim POS_TOP As Integer
Dim POS_LEFT As Integer
Dim POS_TAB As Integer
Dim LINEA As Boolean
Dim nn As Integer
Dim X1, Y1, X2, Y2 As Integer
Dim llave_rep01 As rdoResultset
Dim PS_REP01 As rdoQuery


Public Sub LLENA_CAMPOS1()
Static UNICO As String
Dim TEXTO As TextBox
Dim TOP1 As Integer
Dim LEFT1 As Integer
Dim M As Integer

Dim CODI  As String * 3
M = 0
nn = 2
Do Until Val(tra_llave(nn)) = 0 Or nn = 62
NUMERO = tra_llave(nn)
frmDiseno.cheactivo(M).Visible = True
frmDiseno.OBJETO(M).Visible = True
CODI = Format(tra_llave(nn), "000")
frmDiseno.OBJETO(M).text = "WF" + CODI & " - " & Trim(tra_llave(nn + 1))
frmDiseno.OBJETO(M).Tag = tra_llave(nn)
M = M + 1
nn = nn + 4
Loop
'NUMERO = tra_llave(2)
pub_cadena = "SELECT * FROM REP_TRANSA WHERE REP_CODTRA = ? ORDER BY REP_CODTRA"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = Val(trans.text)
llave_rep01.Requery
If llave_rep01.EOF Then
  MsgBox " Favor de Definir el Reporte. (con la opcion de Graba) ", 48, Pub_Titulo
Else
 M = -1
 For fila = 2 To 17
   M = M + 1
   If Trim(Nulo_Valors(llave_rep01(fila))) <> "" Then
     frmDiseno.cheactivo(M).Value = 1
   End If
 Next fila
 If Nulo_Valors(llave_rep01!REP_ACTIVO) = "A" Then cherepactivo.Value = 1
 If Nulo_Valors(llave_rep01!REP_IMP) = "A" Then cheimp.Value = 1
 If Trim(Nulo_Valors(llave_rep01!REP_CIA)) <> "" Then checia.Value = 1
 If Trim(Nulo_Valors(llave_rep01!REP_DIA)) <> "" Then chedia.Value = 1
 If Trim(Nulo_Valors(llave_rep01!REP_USU)) <> "" Then cheusu.Value = 1
 If Trim(Nulo_Valors(llave_rep01!REP_TRA)) <> "" Then chetra.Value = 1
 txtruta.text = Trim(Nulo_Valors(llave_rep01!REP_RUTA))
 txtnombre.text = Trim(Nulo_Valors(llave_rep01!REP_TRA_NOMBRE))
 txtdirec.text = Trim(Nulo_Valors(llave_rep01!REP_TRA_DIREC))
 txtruc.text = Trim(Nulo_Valors(llave_rep01!REP_TRA_RUC))
 txtplaca.text = Trim(Nulo_Valors(llave_rep01!REP_TRA_PLACA))
End If


End Sub

Private Sub cmdExaminar_Click()
Load frmExaminar
frmExaminar.Show 1
End Sub

Private Sub cmdGene_Click()
''im dd As Excel.Application
Dim wsfile
Dim xl As Object
Dim xcuenta As Integer
If Trim(txtruta.text) = "" Then
 MsgBox "Ingrese Una Ruta para el Archivo.", Pub_Titulo
 Exit Sub
End If
If Trim(Dir(txtruta.text, vbArchive)) <> "" Then
  pub_mensaje = "El Archivo Existe, Desea Eliminarlo y Volverlo a Generar diga: < Si > , Desea Abrir al Archivo para modificaciones diga: < No > .. ?"
  Pub_Respuesta = MsgBox(pub_mensaje, vbDefaultButton2 + vbQuestion + vbYesNoCancel, Pub_Titulo)
  If Pub_Respuesta = vbCancel Then
    Exit Sub
  ElseIf Pub_Respuesta = vbNo Then
     If xl Is Nothing Then
        Set xl = CreateObject("Excel.Application")
        xl.Workbooks.Open Trim(txtruta.text)
        xl.Application.Visible = True
        Set xl = Nothing
     End If
     Exit Sub
  End If
  On Error GoTo OJO:
   Kill Trim(txtruta.text)
  On Error GoTo 0
End If
wsfile = Trim(txtruta.text)
If xl Is Nothing Then
  Set xl = CreateObject("Excel.Application")
End If

xl.SheetsInNewWorkbook = 1
xl.Workbooks.Add
xl.Worksheets(1).Name = "REPORTE"
xl.Windows(1).Caption = Trim(trans.text) + " - " + Trim(transa.Caption)
xl.Windows(1).WindowState = xlMaximized
xl.Workbooks(1).SaveAs wsfile
xcuenta = 0
For fila = 0 To OBJETO().Count - 1
  If cheactivo(fila).Value = 1 Then
     xcuenta = xcuenta + 1
     xl.Cells(xcuenta, 1) = Left(OBJETO(fila).text, 5)
     xl.Cells(xcuenta, 1).Name = Left(OBJETO(fila).text, 5)
  End If
Next fila
xcuenta = xcuenta + 1
If checia.Value = 1 Then
  xl.Cells(xcuenta, 1) = "WF600"
  xl.Cells(xcuenta, 1).Name = "WF600"
End If
xcuenta = xcuenta + 1
If chedia.Value = 1 Then
  xl.Cells(xcuenta, 1) = "WF601"
  xl.Cells(xcuenta, 1).Name = "WF601"
End If
xcuenta = xcuenta + 1
If cheusu.Value = 1 Then
  xl.Cells(xcuenta, 1) = "WF602"
  xl.Cells(xcuenta, 1).Name = "WF602"
End If
xcuenta = xcuenta + 1
If chetra.Value = 1 Then
  xl.Cells(xcuenta, 1) = "WF603"
  xl.Cells(xcuenta, 1).Name = "WF603"
End If

xl.Application.Visible = True
Set xl = Nothing
Exit Sub
OJO:
If Err.Number = 70 Then
  MsgBox "Hoja de Calculo esta Abirta , debe cerrarlo para continuar", 48, Pub_Titulo
Else
  Resume Next
End If
End Sub

Private Sub Command2_Click()
Unload frmDiseno
'FRMTRANS.Show
End Sub

Private Sub Form_Load()
Dim a As Integer
LIMPIA
BLOQUEA
trans.MaxLength = 4
End Sub



Private Sub grabar_Click()
Dim M, N As Integer
Dim msg, estilo, respuesta As String
If Trim(txtruta.text) = "" And txtruta.Enabled Then
  MsgBox "Debe especificar una ruta ", 48, Pub_Titulo
  txtruta.SetFocus
  Exit Sub
End If
If llave_rep01.EOF Then
 llave_rep01.AddNew
 llave_rep01!REP_CODTRA = Val(trans.text)
Else
 llave_rep01.Edit
End If
 M = -1
 For fila = 2 To 17
   M = M + 1
   llave_rep01(fila) = " "
   If frmDiseno.cheactivo(M).Value = 1 Then
    llave_rep01(fila) = Left(OBJETO(M).text, 5)
   End If
 Next fila
 llave_rep01!REP_ACTIVO = " "
 If cherepactivo.Value = 1 Then llave_rep01!REP_ACTIVO = "A"
 llave_rep01!REP_IMP = " "
 If cheimp.Value = 1 Then llave_rep01!REP_IMP = "A"
 llave_rep01!REP_CIA = " "
 If checia.Value = 1 Then llave_rep01!REP_CIA = "WF600"
 llave_rep01!REP_DIA = " "
 If chedia.Value = 1 Then llave_rep01!REP_DIA = "WF601"
 llave_rep01!REP_USU = " "
 If cheusu.Value = 1 Then llave_rep01!REP_USU = "WF602"
 llave_rep01!REP_TRA = " "
 If chetra.Value = 1 Then llave_rep01!REP_TRA = "WF603"
 llave_rep01!REP_RUTA = txtruta.text
 llave_rep01!REP_TRA_NOMBRE = Trim(txtnombre.text)
 llave_rep01!REP_TRA_DIREC = Trim(txtdirec.text)
 llave_rep01!REP_TRA_RUC = Trim(txtruc.text)
 llave_rep01!REP_TRA_PLACA = Trim(txtplaca.text)
 
 llave_rep01.Update
LIMPIA
BLOQUEA
Azul trans, trans
End Sub

Private Sub listatra_KeyPress(KeyAscii As Integer)
If Trim(listatra.text) <> "" And KeyAscii = 13 Then
  trans.text = Left(listatra.text, 4)
 Trans_KeyPress 13
 listatra.Visible = False
End If
End Sub

Private Sub Trans_KeyPress(KeyAscii As Integer)
Dim tf As String
Dim SN As String
Dim Control As Object
Dim a As Integer
SOLO_ENTERO KeyAscii
If KeyAscii = 27 Then
 Unload frmDiseno
 Exit Sub
End If
If KeyAscii = 13 Then
  If trans.text = "" Then
    trans.SetFocus
    Exit Sub
  End If
End If
If KeyAscii <> 13 Then
   GoTo SALIR
End If


'f = frmDiseno.trans.text
'If f = False Then
'   MsgBox "DATO DEBE SER NUMERICO ", 64
'   GoTo SALIR
'End If
PUB_CODTRA = frmDiseno.trans.text
LK_CODTRA = PUB_CODTRA
If frmDiseno.trans.text <> "" Then
   SQ_OPER = 1
   LEER_TRA_LLAVE
   If tra_llave.EOF Then
      MsgBox "TRANSACCION NO EXISTE...", 48, Pub_Titulo
      Azul frmDiseno.trans, frmDiseno.trans
      GoTo SALIR
   Else
      LK_NOMTRA = tra_llave(1)
      LK_CODTRA = tra_llave(0)
      frmDiseno.transa.Caption = LK_NOMTRA
      LIMPIA
      LLENA_CAMPOS1
      DESBLOQUEA
      fop2.Enabled = True
      If LK_CODTRA = 2401 Then
        fop2.Enabled = False
        cheusu.Enabled = False
        chetra.Enabled = False
        checia.Enabled = False
        chedia.Enabled = False
        txtruta.Enabled = False
        cmdGene.Enabled = False
        cmdExaminar.Enabled = False
      End If
  
   End If
End If

SALIR:

End Sub


Private Sub trans_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
 LLENA_LISTRANSA frmDiseno.listatra
 frmDiseno.listatra.Left = 1980
 frmDiseno.listatra.Top = 0
 frmDiseno.listatra.Width = 4500
 frmDiseno.listatra.Height = 3000
 frmDiseno.listatra.Visible = True
 frmDiseno.listatra.SetFocus
End If
End Sub

Public Sub LIMPIA()
For fila = 0 To OBJETO.Count - 1
  OBJETO(fila).Visible = False
  frmDiseno.cheactivo(fila).Visible = False
  frmDiseno.cheactivo(fila).Value = 0
Next fila
cherepactivo.Value = 0
frmDiseno.checia.Value = 0
frmDiseno.chedia.Value = 0
frmDiseno.cheimp.Value = 0
frmDiseno.chetra.Value = 0
frmDiseno.cheusu.Value = 0
frmDiseno.txtruta.text = ""
frmDiseno.txtnombre.text = ""
frmDiseno.txtdirec.text = ""
frmDiseno.txtruc.text = ""
frmDiseno.txtplaca.text = ""
End Sub
Public Sub BLOQUEA()
fop1.Enabled = False
fop3.Enabled = False
cherepactivo.Enabled = False
frmDiseno.checia.Enabled = False
frmDiseno.chedia.Enabled = False
frmDiseno.cheimp.Enabled = False
frmDiseno.chetra.Enabled = False
frmDiseno.cheusu.Enabled = False
frmDiseno.txtruta.Enabled = False
Grabar.Enabled = False
End Sub
Public Sub DESBLOQUEA()
Grabar.Enabled = True
fop1.Enabled = True
fop3.Enabled = True
cherepactivo.Enabled = True
frmDiseno.checia.Enabled = True
frmDiseno.chedia.Enabled = True
frmDiseno.cheimp.Enabled = True
frmDiseno.chetra.Enabled = True
frmDiseno.cheusu.Enabled = True
frmDiseno.txtruta.Enabled = True
End Sub

