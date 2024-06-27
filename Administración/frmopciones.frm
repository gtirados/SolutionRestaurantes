VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmopciones 
   Caption         =   "Opciones de Proveedor"
   ClientHeight    =   5370
   ClientLeft      =   525
   ClientTop       =   1560
   ClientWidth     =   8685
   Icon            =   "frmopciones.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   8685
   Begin VB.CheckBox CRE 
      Caption         =   "CREDITO"
      Height          =   255
      Left            =   720
      TabIndex        =   37
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cambiar 
      Caption         =   "Cambiar"
      Height          =   255
      Left            =   4680
      TabIndex        =   36
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox numfac2 
      Height          =   285
      Left            =   3480
      TabIndex        =   34
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox ser2 
      Height          =   285
      Left            =   2880
      TabIndex        =   33
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox numfac 
      Height          =   285
      Left            =   1440
      TabIndex        =   32
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox ser 
      Height          =   285
      Left            =   840
      TabIndex        =   31
      Top             =   3840
      Width           =   375
   End
   Begin VB.ComboBox fbg 
      Height          =   315
      ItemData        =   "frmopciones.frx":0442
      Left            =   240
      List            =   "frmopciones.frx":0452
      TabIndex        =   30
      Text            =   "Combo1"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "COPIAR TAB_TIPREG "
      Height          =   315
      Left            =   5760
      TabIndex        =   29
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Frame Frame3 
      Height          =   3735
      Left            =   5640
      TabIndex        =   17
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton Command6 
         Caption         =   "COMAESTcopiar a OTRA CIA"
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   2160
         Width           =   2655
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Traer datos"
         Height          =   315
         Left            =   1440
         TabIndex        =   27
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Pasar datos"
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txtchar 
         Height          =   285
         Left            =   720
         MaxLength       =   2
         TabIndex        =   23
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtcia 
         Height          =   285
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   21
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtfami 
         Height          =   285
         Left            =   720
         TabIndex        =   20
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Procesar"
         Height          =   315
         Left            =   600
         TabIndex        =   18
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Pasa deTRA_KEY >= 800 a Excel (TRANSA.XLS) "
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   26
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "CHAR"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "cia:"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   22
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Copiar  articulos x cod. familia"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Copiar Datos a Otra Compañia"
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5535
      Begin VB.ComboBox cmdcias 
         Height          =   315
         ItemData        =   "frmopciones.frx":0462
         Left            =   1440
         List            =   "frmopciones.frx":0472
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Procesar"
         Height          =   315
         Left            =   4080
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox CP 
         Height          =   315
         ItemData        =   "frmopciones.frx":049F
         Left            =   4560
         List            =   "frmopciones.frx":04A9
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmdtablas 
         Height          =   315
         ItemData        =   "frmopciones.frx":04C2
         Left            =   1440
         List            =   "frmopciones.frx":04D2
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "de Cia. Actual.       a esta Cia."
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Cli. /Prov :"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblCia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Tablas a Copiar:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones  para el Formulario de Clientes :"
      Height          =   2295
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   5535
      Begin VB.CommandButton cmdguardar 
         Caption         =   "Guardar Cambios"
         Height          =   495
         Left            =   4440
         TabIndex        =   16
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ver"
         Height          =   375
         Left            =   4440
         TabIndex        =   15
         Top             =   600
         Width           =   975
      End
      Begin VB.ListBox listACCESO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   14
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Cargar el Formulario de clientes para Configurar.(Visible / Invisible)"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "Ce&rrar"
      Height          =   375
      Left            =   6960
      TabIndex        =   11
      Top             =   4680
      Width           =   1335
   End
   Begin ComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Label Label4 
      Caption         =   "por"
      Height          =   255
      Left            =   2520
      TabIndex        =   35
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label lblmensa 
      Caption         =   "Procensando  . . . "
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   3975
   End
End
Attribute VB_Name = "frmopciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xl As Object
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset
Dim PS_REP02 As rdoQuery
Dim llave_rep02 As rdoResultset
Dim PS_REP03 As rdoQuery
Dim llave_rep03 As rdoResultset
Dim PS_REP04 As rdoQuery
Dim llave_rep04 As rdoResultset
Dim PS_REP05 As rdoQuery
Dim llave_rep05 As rdoResultset

Private Sub cambiar_Click()

pub_cadena = "SELECT FAR_NUMFAC, FAR_NUMSER FROM FACART WHERE FAR_CODCIA = ? AND FAR_FBG = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ? AND FAR_ESTADO <> 'E' "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = Trim(fbg.text)
PS_REP01(2) = ser.text
PS_REP01(3) = numfac.text
llave_rep01.Requery
If llave_rep01.EOF Then
   MsgBox "Verificar numeracion ", 48, Pub_Titulo
   Exit Sub
End If

pub_cadena = "SELECT ALL_NUMFAC, ALL_NUMSER FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FBG = ? AND ALL_NUMSER = ? AND ALL_NUMFAC = ? AND ALL_FLAG_EXT <> 'E'  "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP02(0) = LK_CODCIA
PS_REP02(1) = Trim(fbg.text)
PS_REP02(2) = ser.text
PS_REP02(3) = numfac.text
llave_rep02.Requery
If llave_rep02.EOF Then
   MsgBox "Verificar numeracion ", 48, Pub_Titulo
   Exit Sub
End If


pub_cadena = "SELECT CAR_NUMFAC, CAR_NUMSER FROM CARTERA WHERE CAR_CODCIA = ? AND CAR_FBG = ? AND CAR_NUMSER = ? AND CAR_NUMFAC = ?  "
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP03(0) = LK_CODCIA
PS_REP03(1) = Trim(fbg.text)
PS_REP03(2) = ser.text
PS_REP03(3) = numfac.text
llave_rep03.Requery
If llave_rep03.EOF Then
   MsgBox "Verificar numeracion ", 48, Pub_Titulo
   Exit Sub
End If


pub_cadena = "SELECT CAA_NUMFAC, CAA_NUMSER FROM CARACU WHERE CAA_CODCIA = ? AND CAA_FBG = ? AND CAA_NUMSER = ? AND CAA_NUMFAC = ?  "
Set PS_REP04 = CN.CreateQuery("", pub_cadena)
Set llave_rep04 = PS_REP04.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP04(0) = LK_CODCIA
PS_REP04(1) = Trim(fbg.text)
PS_REP04(2) = ser.text
PS_REP04(3) = numfac.text
llave_rep04.Requery
If llave_rep04.EOF Then
   MsgBox "Verificar numeracion ", 48, Pub_Titulo
   Exit Sub
End If

pub_mensaje = " ¿Desea Continuar... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If

Do Until llave_rep01.EOF
    llave_rep01.Edit
    llave_rep01!far_NUMSER = ser2.text
    llave_rep01!far_NUMFAC = Val(numfac2.text)
    llave_rep01.Update
llave_rep01.MoveNext
Loop

Do Until llave_rep02.EOF
    llave_rep02.Edit
    llave_rep02!ALL_NUMSER = ser2.text
    llave_rep02!ALL_NUMFAC = Val(numfac2.text)
    llave_rep02.Update
llave_rep02.MoveNext
Loop

If CRE.Value <> 1 Then GoTo FIN

Do Until llave_rep03.EOF
llave_rep03.Edit
llave_rep03!CAR_NUMSER = ser2.text
llave_rep03!CAR_NUMFAC = Val(numfac2.text)
llave_rep03.Update
llave_rep03.MoveNext
Loop

Do Until llave_rep04.EOF
llave_rep04.Edit
llave_rep04!CAA_NUMSER = ser2.text
llave_rep04!CAA_NUMFAC = Val(numfac2.text)
llave_rep04.Update
llave_rep04.MoveNext
Loop

FIN:
MsgBox "CAMBIO TERMINADO"
End Sub

Private Sub cmdcerrar_Click()
 Unload frmopciones
End Sub

Private Sub cmdguardar_Click()
If listACCESO.ListCount = 0 Then
 MsgBox " No Procede ... revisar, debe ver al menos uno en la lista", 48, Pub_Titulo
 Exit Sub
End If
Dim wcambio As String
wcambio = ""
For i = 0 To listACCESO.ListCount - 1
  listACCESO.ListIndex = i
  If listACCESO.Selected(i) Then
      wcambio = wcambio & "." & Format(Val(Trim(Right(listACCESO.text, 4))), "00")
  End If
Next i
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_PAR_LLAVE
If wcambio = "" Then
 wcambio = " "
End If
par_llave.Edit
par_llave!par_acceso_cli = wcambio
par_llave.Update
MsgBox "Actualizado. Debe Reuniciar la Compañia para ver los cambios.", 48, Pub_Titulo
End Sub

Private Sub cmdtablas_Click()

If cmdtablas.text = "CLIENTES" Then
  CP.Visible = True
  CP.ListIndex = 0
Else
  CP.Visible = False
End If
End Sub

Private Sub Command1_Click()

If cmdtablas.text = "" Then
 MsgBox "seleciona tabla ."
 Exit Sub
End If
If cmdtablas.text = "CLIENTES" Then
  If CP.text = "" Then
    MsgBox "seleciona tabla ."
    Exit Sub
  End If
End If
If cmdcias.text = "" Then
 MsgBox "seleciona campañia."
 Exit Sub
End If
If cmdtablas.text = "" Then
 MsgBox "seleciona tabla ."
 Exit Sub
End If
If Trim(Left(cmdcias.text, 2)) = Trim(lblCia.Caption) Then
MsgBox "NO se puede sobrescribir a si mismo..."
Exit Sub
End If
PB.Value = 0
PB.Min = 0
PB.Visible = True
lblmensa.Visible = True
DoEvents
PB.Visible = True
DoEvents
Dim tabla As String
Dim SELE As String
Dim WCAMPO As String
SELE2 = ""
tabla = cmdtablas.text
WCAMPO = Left(tabla, 3)
If tabla = "ARTICULO" Then
 WCAMPO = "ARM"
ElseIf tabla = "CARACU" Then
 WCAMPO = "CAA"
ElseIf tabla = "COMOV" Then
 WCAMPO = "COV"
ElseIf tabla = "FACART" Then
 WCAMPO = "FAR"
ElseIf tabla = "CONTABILIDAD" Then
 WCAMPO = "CNT"
ElseIf tabla = "CLIENTES" Then
 SELE2 = " AND CLI_CP = '" + Left(CP.text, 1) + "'"
ElseIf tabla = "CALENDARIO" Then
 WCAMPO = "CAL"
ElseIf tabla = "CONCEPTO" Then
 WCAMPO = "CON"
ElseIf tabla = "PLANTILLA" Then
 WCAMPO = "PLT"
ElseIf tabla = "COPARAM" Then
 WCAMPO = "COP"
End If

SELE = " WHERE " + WCAMPO + "_" + "CODCIA = '" + lblCia.Caption + "'" + SELE2
pub_cadena = "SELECT * FROM " + tabla + SELE
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

SELE = " WHERE " + WCAMPO + "_" + "CODCIA = '" + Left(cmdcias.text, 2) + "'" + SELE2
pub_cadena = "SELECT * FROM " + tabla + SELE
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurValues)

llave_rep01.Requery

llave_rep02.Requery
wsflag = ""
If Not llave_rep02.EOF Then
    pub_mensaje = " !!! Hay Datos en esta compañia...   ¿Desea Sobrescribir... ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
       PB.Visible = False
       lblmensa.Visible = False
       Exit Sub
    End If
    CN.Execute "Begin Transaction", rdExecDirect
    SELE = " WHERE " + WCAMPO + "_" + "CODCIA = '" + Left(cmdcias.text, 2) + "'" + SELE2
    pub_cadena = "DELETE " + tabla + SELE
    CN.Execute pub_cadena, rdExecDirect
    wsflag = "A"
End If
On Error GoTo ESCAPA
If wsflag <> "A" Then
 CN.Execute "Begin Transaction", rdExecDirect
End If
PB.Max = llave_rep01.RowCount
llave_rep01.MoveFirst
Do Until llave_rep01.EOF
 PB.Value = PB.Value + 1
 llave_rep02.AddNew
 For fila = 0 To llave_rep01.rdoColumns.Count - 1
   If fila = 1 Then
     If tabla = "COMOV" Or tabla = "PRECIOS" Or tabla = "CARACU" Or tabla = "CARTERA" Or tabla = "COMAEST" Or tabla = "CONTABILIDAD" Or tabla = "CALENDARIO" Or tabla = "CONCEPTOS" Or tabla = "PLANTILLA" Or tabla = "COPARAM" Then
       If Not IsNull(llave_rep01.rdoColumns(fila)) Then llave_rep02.rdoColumns(fila) = llave_rep01.rdoColumns(fila)
     Else
       llave_rep02.rdoColumns(fila) = Left(cmdcias.text, 2)
     End If
   Else
     If Not IsNull(llave_rep01.rdoColumns(fila)) Then llave_rep02.rdoColumns(fila) = llave_rep01.rdoColumns(fila)
   End If
 Next fila
 If tabla = "COMAEST" Or tabla = "CONTABILIDAD" Or tabla = "CONCEPTOS" Or tabla = "PLANTILLA" Or tabla = "COPARAM" Then
   llave_rep02.rdoColumns(0) = Left(cmdcias.text, 2)
 End If
 If tabla = "PRECIOS" Then
   llave_rep02.rdoColumns(0) = Left(cmdcias.text, 2)
 End If
 If tabla = "CARACU" Then
   llave_rep02.rdoColumns(2) = Left(cmdcias.text, 2)
 End If
 If tabla = "CARTERA" Then
   llave_rep02.rdoColumns(2) = Left(cmdcias.text, 2)
 End If
 If tabla = "COMOV" Or tabla = "CALENDARIO" Then
   llave_rep02.rdoColumns(0) = Left(cmdcias.text, 2)
 End If
 llave_rep02.Update
 
 llave_rep01.MoveNext
Loop
CN.Execute "Commit Transaction", rdExecDirect
MsgBox "OPERACION COMPLETADA . . . "
PB.Visible = False
lblmensa.Visible = False
On Error GoTo 0
Exit Sub
ESCAPA:
MsgBox Err.Description
'Resume Next
CN.Execute "Rollback Transaction", rdExecDirect
End Sub

Private Sub Command2_Click()
LLENAR_CLI_ACCESO
cmdguardar.Visible = True
End Sub

Private Sub Command3_Click()

Dim ORIGINAL  As Currency
Dim ALTERNO  As String
Dim PS1 As rdoQuery
Dim ps1_articulos As rdoResultset
Dim PS2 As rdoQuery
Dim ps2_articulos As rdoResultset

Dim PS1_PRE As rdoQuery
Dim ps1_precios As rdoResultset
Dim PS2_PRE As rdoQuery
Dim ps2_precios As rdoResultset

If Trim(txtcia.text) = "" Then
 MsgBox "cia"
 Exit Sub
End If
If Trim(txtfami.text) = "" Then
 MsgBox "cia"
 Exit Sub
End If

pub_cadena = "SELECT * FROM ARTICULO WHERE ARM_CODCIA = ? AND ARM_CODART = ? "
Set PS_1 = CN.CreateQuery("", pub_cadena)
PS_1(0) = 0
PS_1(1) = 0
Set ps1_articulos = PS_1.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM ARTICULO "
Set PS2 = CN.CreateQuery("", pub_cadena)
Set ps2_articulos = PS2.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = ? AND PRE_CODART = ? "
Set PS1_PRE = CN.CreateQuery("", pub_cadena)
PS1_PRE(0) = 0
PS1_PRE(1) = 0
Set ps1_precios = PS1_PRE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM PRECIOS"
Set PS2_PRE = CN.CreateQuery("", pub_cadena)
Set ps2_precios = PS2_PRE.OpenResultset(rdOpenKeyset, rdConcurValues)


'ARTI
pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ?"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = LK_CODCIA
'PS_REP01(1) = Val(txtfami.text)

pub_cadena = "SELECT * FROM  ARTI WHERE ART_CODCIA = '" & txtcia.text & "'"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurValues)
llave_rep01.Requery
llave_rep02.Requery
wsflag = ""
PB.Visible = True
PB.Value = 0
PB.Max = llave_rep01.RowCount
PB.Min = 0

If Not llave_rep02.EOF Then
    pub_mensaje = " !!! Hay Datos en esta compañia...   ¿Desea agregar mas articulos. ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
       PB.Visible = False
       lblmensa.Visible = False
       Exit Sub
    End If

'  Exit Sub
End If
On Error GoTo ESCAPA
CN.Execute "Begin Transaction", rdExecDirect
PB.Max = llave_rep01.RowCount
llave_rep01.MoveFirst
Do Until llave_rep01.EOF
 PB.Value = PB.Value + 1
 ORIGINAL = GENERA_CODI(txtcia.text)
 ALTERNO = Trim(llave_rep01.rdoColumns(27)) '& Trim(txtchar.text)
 llave_rep02.AddNew
 For fila = 0 To llave_rep01.rdoColumns.Count - 1
    If fila = 0 Then
      llave_rep02.rdoColumns(fila) = ORIGINAL
    ElseIf fila = 1 Then
      llave_rep02.rdoColumns(fila) = Trim(txtcia.text)
    ElseIf fila = 27 Then
      llave_rep02.rdoColumns(fila) = ALTERNO
    Else
      If Not IsNull(llave_rep01.rdoColumns(fila)) Then
         llave_rep02.rdoColumns(fila) = llave_rep01.rdoColumns(fila)
      End If
    End If
 Next fila
 
 GoSub GRABA_ARTICULOS
 GoSub GRABA_PRECIOS
 llave_rep02.Update
 llave_rep01.MoveNext
Loop
CN.Execute "Commit Transaction", rdExecDirect
MsgBox "OPERACION COMPLETADA . . . "
PB.Visible = False
Exit Sub
GRABA_ARTICULOS:
PS_1(0) = LK_CODCIA
PS_1(1) = llave_rep01.rdoColumns(0)
ps1_articulos.Requery
ps2_articulos.Requery
If Not ps2_articulos.EOF Then
'  MsgBox "HAY DATOS EN LA CIA QUE DESEAE COPIAR,  -  BORRAR EN  : ARTI, ARTICULOS, PRECIOS"
'  Exit Sub
End If
ps1_articulos.MoveFirst
Do Until ps1_articulos.EOF
 ps2_articulos.AddNew
 For fila = 0 To ps1_articulos.rdoColumns.Count - 1
    If fila = 0 Then
      ps2_articulos.rdoColumns(fila) = ORIGINAL
    ElseIf fila = 1 Then
      ps2_articulos.rdoColumns(fila) = Trim(txtcia.text)
    Else
      ps2_articulos.rdoColumns(fila) = Nulo_Valor0(ps1_articulos.rdoColumns(fila))
    End If
 Next fila
 ps2_articulos.Update
 ps1_articulos.MoveNext
Loop
Return

GRABA_PRECIOS:

PS1_PRE(0) = LK_CODCIA
PS1_PRE(1) = llave_rep01.rdoColumns(0)
ps1_precios.Requery
ps2_precios.Requery

ps1_precios.MoveFirst
Do Until ps1_precios.EOF
 ps2_precios.AddNew
 For fila = 0 To ps1_precios.rdoColumns.Count - 1
    If fila = 1 Then
      ps2_precios.rdoColumns(fila) = ORIGINAL
    ElseIf fila = 0 Then
      ps2_precios.rdoColumns(fila) = Trim(txtcia.text)
    Else
      ps2_precios.rdoColumns(fila) = Nulo_Valor0(ps1_precios.rdoColumns(fila))
    End If
 Next fila
 ps2_precios.Update
 ps1_precios.MoveNext
Loop

Return

Exit Sub
ESCAPA:
MsgBox Err.Description
Resume Next

CN.Execute "Rollback Transaction", rdExecDirect
'Resume Next
End Sub

Private Sub Command4_Click()
Dim ORIGINAL  As Currency
Dim ALTERNO  As String
Dim PS1 As rdoQuery
Dim ps1_articulos As rdoResultset
Dim PS2 As rdoQuery
Dim ps2_articulos As rdoResultset

Dim PS1_PRE As rdoQuery
Dim ps1_precios As rdoResultset
Dim PS2_PRE As rdoQuery
Dim ps2_precios As rdoResultset

pub_cadena = "SELECT * FROM  TRANSACCION WHERE TRA_KEY >= ? ORDER BY TRA_KEY"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = 8000

llave_rep01.Requery
wsflag = ""
PB.Visible = True
PB.Value = 0
PB.Max = llave_rep01.RowCount
PB.Min = 0

If llave_rep01.EOF Then
   MsgBox " NO HAY DATOS DE TRANSACCION ", 48, Pub_Titulo
    PB.Visible = False
    lblmensa.Visible = False
    Exit Sub
End If
GoSub WEXCEL
PB.Max = llave_rep01.RowCount
WR = "A2:DT" & llave_rep01.RowCount
xl.Range(WR).Delete
xcuenta = 0
'xl.Application.Visible = True
Do Until llave_rep01.EOF
 xcuenta = xcuenta + 1
 PB.Value = PB.Value + 1
 For fila = 0 To llave_rep01.rdoColumns.Count - 1
   xl.Cells(xcuenta + 1, fila + 1) = llave_rep01.rdoColumns(fila)

 Next fila
 llave_rep01.MoveNext
Loop
  xl.Application.Visible = True
  PB.Visible = False
  lblmensa.Visible = False
  Set xl = Nothing

MsgBox "OPERACION COMPLETADA . . . "
PB.Visible = False
Exit Sub
WEXCEL:
  Dim dd As Excel.Application
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  'Dim RR As Excel.Application
  'RR.Workbooks.Open " dwfSDF",
  xl.Workbooks.Open "C:\ADMIN\STANDAR\TRANSA.xls", False
Return

End Sub

Private Sub Command5_Click()
Dim ORIGINAL  As Currency
Dim ALTERNO  As String
Dim PS1 As rdoQuery
Dim ps1_articulos As rdoResultset
Dim PS2 As rdoQuery
Dim ps2_articulos As rdoResultset

Dim PS1_PRE As rdoQuery
Dim ps1_precios As rdoResultset
Dim PS2_PRE As rdoQuery
Dim ps2_precios As rdoResultset

pub_cadena = "SELECT * FROM  TRANSACCION WHERE TRA_KEY >= ? ORDER BY TRA_KEY"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0

Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = 8000
llave_rep01.Requery

wsflag = ""
PB.Visible = True
PB.Value = 0
'PB.Max = llave_rep01.RowCount
'PB.Min = 0
'If llave_rep01.EOF Then
'   MsgBox " NO HAY DATOS DE TRANSACCION ", 48, Pub_Titulo
'    PB.Visible = False
'    lblmensa.Visible = False
'    Exit Sub
'End If
GoSub WEXCEL
lblmensa.Visible = True
DoEvents
'PB.Max = llave_rep01.RowCount
xcuenta = 0
WS = "N"
Do Until WS = "S"
 xcuenta = xcuenta + 1
 lblmensa.Caption = "Un Momento.." + wcuenta
 DoEvents
 If Val(xl.Cells(xcuenta + 1, 1)) = 0 Then
   Exit Do
 End If
 llave_rep01.AddNew
 For fila = 0 To llave_rep01.rdoColumns.Count - 1
    llave_rep01.rdoColumns(fila) = xl.Cells(xcuenta + 1, fila + 1)
 Next fila
 llave_rep01.Update
Loop
  PB.Visible = False
  lblmensa.Visible = False
  xl.Application.Visible = True
  Set xl = Nothing
MsgBox "OPERACION COMPLETADA . . . "
PB.Visible = False
Exit Sub
WEXCEL:
  Dim dd As Excel.Application
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  'Dim RR As Excel.Application
  'RR.Workbooks.Open " dwfSDF",
  On Error GoTo Ver
  xl.Workbooks.Open "C:\ADMIN\STANDAR\TRANSA.xls", False
  On Error GoTo 0
Return
Exit Sub
Ver:
If Err.Number = 1004 Then
  MsgBox "Crear al Pase de Informacion de Transacciones.."
  
End If
Set xl = Nothing
End Sub

Private Sub Command6_Click()
Dim ORIGINAL  As Currency
Dim ALTERNO  As String
Dim PS1 As rdoQuery
Dim ps1_articulos As rdoResultset
Dim PS2 As rdoQuery
Dim ps2_articulos As rdoResultset

Dim PS1_PRE As rdoQuery
Dim ps1_precios As rdoResultset
Dim PS2_PRE As rdoQuery
Dim ps2_precios As rdoResultset
Dim MIO As String
If Trim(txtcia.text) = "" Then
' MsgBox "cia"
' Exit Sub
End If
If Trim(txtfami.text) = "" Then
' MsgBox "cia"
' Exit Sub
End If


pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ? AND ART_FAMILIA = 28 "
Set PS_1 = CN.CreateQuery("", pub_cadena)
PS_1(0) = 0
Set ps1_articulos = PS_1.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODART = ? AND PRE_CODCIA = ? "
Set PS2 = CN.CreateQuery("", pub_cadena)
PS2(0) = 0
PS2(1) = 0
Set ps2_articulos = PS2.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_1.rdoParameters(0) = LK_CODCIA
ps1_articulos.Requery
MsgBox ps1_articulos.RowCount
MsgBox ps1_articulos.AbsolutePosition
'revisar
Do Until ps1_articulos.EOF
   PS2.rdoParameters(0) = ps1_articulos!ART_KEY
   PS2.rdoParameters(1) = LK_CODCIA
   ps2_articulos.Requery
   'If ps2_articulos.RowCount > 1 Then Stop
   
   ps2_articulos.Edit
   ps2_articulos!pre_pre5 = ps2_articulos!pre_pre1 - (ps2_articulos!pre_pre1 * 0.1)
   
   ps2_articulos.Update
   ps1_articulos.MoveNext
Loop
MsgBox "PROCESO TERMINADO", 48, Pub_Titulo


Exit Sub
Exit Sub


pub_cadena = "SELECT DISTINCT FAR_CODART FROM FACART WHERE FAR_CODCIA = '" & LK_CODCIA & "' AND FAR_CODART <> 0 ORDER BY FAR_CODART"
Set PS_1 = CN.CreateQuery("", pub_cadena)
Set ps1_articulos = PS_1.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
ps1_articulos.Requery
pu_codcia = LK_CODCIA
SQ_OPER = 1
Do Until ps1_articulos.EOF
  PUB_KEY = ps1_articulos!far_codart
  LEER_ART_LLAVE
  If art_LLAVE.EOF Then
   MsgBox " VERIFICAR CODIGO...: " & ps1_articulos!far_codart
  End If
  ps1_articulos.MoveNext
Loop
MsgBox "PROCESO CHEQUEADO !!!!"



Exit Sub

Exit Sub
'OJO AQUI ES CONTABILIDAD
pub_cadena = "SELECT * FROM COMAEST WHERE COM_CODCIA = '02'"
Set PS_1 = CN.CreateQuery("", pub_cadena)
Set ps1_articulos = PS_1.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM COMAEST "
Set PS2 = CN.CreateQuery("", pub_cadena)
Set ps2_articulos = PS2.OpenResultset(rdOpenKeyset, rdConcurValues)
'pub_cadena = "DELETE COMOX"
'CN.Execute pub_cadena, rdExecDirect
ps1_articulos.Requery
Do Until ps1_articulos.EOF
 ps2_articulos.AddNew
 For fila = 0 To ps1_articulos.rdoColumns.Count - 1
    ps2_articulos.rdoColumns(fila) = ps1_articulos.rdoColumns(fila)
 Next fila
 ps2_articulos.Update
 ps1_articulos.MoveNext
Loop
MsgBox "PROCESO TERMINADO", 48, Pub_Titulo
End Sub

Private Sub Command7_Click()

Dim ORIGINAL  As Currency
Dim ALTERNO  As String
Dim PS1 As rdoQuery
Dim ps1_articulos As rdoResultset
Dim PS2 As rdoQuery
Dim ps2_articulos As rdoResultset
Dim WCODICA As String * 2
Dim WTAB_TIPREG As Integer
Dim WW
WW = InputBox("TIP_REG PARA COPIAR :")
If WW = "" Then
  Exit Sub
End If
WTAB_TIPREG = Val(WW)
wcodcia = InputBox("A QUE CODCIA COPIAR :")
If wcodcia = "" Then
  Exit Sub
End If

pub_cadena = "SELECT * FROM TABLAS WHERE TAB_TIPREG = ? AND TAB_CODCIA = ? "
Set PS_1 = CN.CreateQuery("", pub_cadena)
PS_1(0) = 0
PS_1(1) = 0
Set ps1_articulos = PS_1.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM TABLAS"
Set PS2 = CN.CreateQuery("", pub_cadena)
Set ps2_articulos = PS2.OpenResultset(rdOpenKeyset, rdConcurValues)

PS_1.rdoParameters(0) = WTAB_TIPREG
PS_1.rdoParameters(1) = LK_CODCIA
ps1_articulos.Requery
ps2_articulos.Requery

'pub_cadena = "DELETE TABLAS WHERE TAB_CODCIA = '" & wcodcia & "' AND TAB_TIPREG = " & WTAB_TIPREG
'CN.Execute pub_cadena, rdExecDirect
Do Until ps1_articulos.EOF
   ps2_articulos.AddNew
   For fila = 0 To ps1_articulos.rdoColumns.Count - 1
     ps2_articulos.rdoColumns(fila) = Nulo_Valor0(ps1_articulos.rdoColumns(fila))
   Next fila
  ps2_articulos.rdoColumns(0) = wcodcia
  ps2_articulos.Update
'  ps2_articulos.CancelUpdate
  ps1_articulos.MoveNext
Loop
MsgBox "PROCESO TERMINADO", 48, Pub_Titulo


End Sub

Private Sub Form_Load()
cmdtablas.Clear
CP.Visible = False
If CN.rdoTables.Count = 0 Then
   MsgBox " Esta bloqueado...", 48, Pub_Titulo
   GoTo FIN
End If
For fila = 0 To CN.rdoTables.Count - 1
    wtab = CN.rdoTables(fila).Name
    If wtab = "DEFCONTXX" Or wtab = "CONTROLL" Or wtab = "CALFER" Or wtab = "ESPESOR" Or wtab = "SUB_TRANSA" Or wtab = "TRANSACCIONL" Or wtab = "USUARIOS" Then
    Else
     cmdtablas.AddItem CN.rdoTables(fila).Name
    End If
Next
cmdcias.Clear
PS_PAR(0) = " "
par.Requery
Do Until par.EOF
 cmdcias.AddItem par!PAR_CODCIA & " - " & par!PAR_NOMBRE
 par.MoveNext
Loop
lblCia.Caption = LK_CODCIA

Exit Sub
FIN:
End Sub


Public Sub LLENAR_CLI_ACCESO()
Static bandera
On Error GoTo SALE
If Not frmCLI.Visible Then
  Load frmCLI
  frmCLI.Show
  Exit Sub
End If

Dim cuenta As Integer
fila = 0
listACCESO.Clear
cuenta = -1
Do Until fila >= frmCLI.Controls.Count
 'MsgBox frmCLI.Controls(fila).Name
 If TypeOf frmCLI.Controls(fila) Is Timer Then
   GoTo OTRITO
 End If
 If TypeOf frmCLI.Controls(fila) Is Label Then
   GoTo OTRITO
 End If
 If frmCLI.Controls(fila).WhatsThisHelpID <> 0 Then
   cuenta = cuenta + 1
   listACCESO.AddItem frmCLI.Controls(fila).Name + String(90, " ") + Str(frmCLI.Controls(fila).WhatsThisHelpID)
   If frmCLI.Controls(fila).Visible Then
     listACCESO.Selected(cuenta) = True
   End If
   'frmCLI.Controls(fila).Visible = True
 End If
 
OTRITO:
 fila = fila + 1
Loop

Exit Sub
SALE:
 MsgBox "Carge el formulario de Clientes .. ", 48, Pub_Titulo

End Sub

Public Function GENERA_CODI(wcia As String) As Double
Dim NUMCAD, FIJO As String
Dim DIGI As String * 2
Dim i, VINT1, VINT2, VINT3, VINT4 As Double
Dim VSTR1, VSTR2, VSTR3, VSTR4 As String
Dim VFIJO As Double
Dim VVARI As Integer
Dim STRpub_cadena As String
Dim INTpub_cadena As Double

PUB_KEY = 0
SQ_OPER = 2
pu_codcia = wcia
LEER_ART_LLAVE

If art_mayor.EOF Then
    NUMCAD = "1"
Else
    art_mayor.MoveLast
    NUMCAD = art_mayor!ART_KEY
End If

VINT2 = 0
NUMCAD = Trim(NUMCAD)
VINT1 = Len(NUMCAD)
If NUMCAD = "1" Or NUMCAD = "2" Or NUMCAD = "0" Then
  VINT2 = 1
  VINT1 = 2
End If
If VINT1 > 1 Then
    VSTR4 = Val(Mid(NUMCAD, 1, VINT1 - 2)) + 1
End If

For i = 1 To VINT1 - 2
   VSTR1 = Mid(VSTR4, i, 1)
   VINT2 = VINT2 + Val(VSTR1)
Next i
VINT3 = VINT2 * 7

VSTR3 = Right(CStr(VINT3), 2)
If Len(VSTR3) = 1 Then
  VSTR3 = "0" & VSTR3
End If
FIJO = VSTR4
STRpub_cadena = FIJO & VSTR3
INTpub_cadena = Val(STRpub_cadena)

GENERA_CODI = INTpub_cadena

End Function

