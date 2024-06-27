VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmDatArti 
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   Icon            =   "FrmDatArti.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox textovar 
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Gridtablas 
      Height          =   4335
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7646
      _Version        =   393216
      Cols            =   5
   End
   Begin MSFlexGridLib.MSFlexGrid grid_comi 
      Height          =   4335
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7646
      _Version        =   393216
      Cols            =   5
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Retornar"
      Height          =   585
      Left            =   2970
      Picture         =   "FrmDatArti.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4845
      Width           =   1020
   End
   Begin VB.CommandButton cmdagregar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Agregar"
      Height          =   585
      Left            =   855
      Picture         =   "FrmDatArti.frx":0C10
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4845
      Width           =   1125
   End
   Begin VB.Label lblbarraos 
      Alignment       =   2  'Center
      BackColor       =   &H008B4914&
      Caption         =   "Solution for Business"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   5520
      Width           =   4935
   End
   Begin VB.Label lblmensa 
      Caption         =   "[F2] = Modificar                                 [DEL] = Eliminar"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   4095
   End
End
Attribute VB_Name = "FrmDatArti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wBAN As Integer
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset

Private Sub cmdAgregar_Click()
Dim wcodciax As String * 2
Dim WNUMTAB As Currency
Dim WCODART As Integer
Dim wvalor
Dim ultimo As Long
Dim Item As Integer
Dim WFLAG
If TEXTOVAR.Visible Then
  MsgBox "Esta Usted. Editanto ....!", 48, Pub_Titulo
  Exit Sub
End If
If PUB_TIPREG = -10 Then GoTo Transportista
If PUB_TIPREG = -5 Then GoTo Graba_Direccion
If cmdAgregar.Caption = "Gragar Cta." Then
  GoTo graba_cta
End If
If Left(cmdAgregar.Caption, 2) = "&G" Then
    Item = 0
    WFLAG = ""
    For fila = 2 To grid_comi.Rows - 1
       If WFLAG = "A" Then
          MsgBox "Verificar , El Ultimo rango debe ser parecido a unico ", 48, Pub_Titulo
          Exit Sub
       End If
       If val(grid_comi.TextMatrix(fila, 1)) <> 0 And val(grid_comi.TextMatrix(fila, 2)) = 0 Then
         WFLAG = "A"
       ElseIf val(grid_comi.TextMatrix(fila, 1)) > val(grid_comi.TextMatrix(fila, 2)) Then
         MsgBox "Verificar , Dia Inicial no debe ser Mayor que la Final ", 48, Pub_Titulo
         Exit Sub
       End If
       If Not val(grid_comi.TextMatrix(fila, 1)) >= Item Then
          MsgBox "Verificar , Dia Inicial  debe estar en Orden ascendente ", 48, Pub_Titulo
          Exit Sub
       End If
       Item = val(grid_comi.TextMatrix(fila, 1))
    Next fila
    On Error GoTo ESCAPA
    CN.Execute "Begin Transaction", rdExecDirect
    SQ_OPER = 2
    LEER_TAB_LLAVE
    Do Until tab_mayor.EOF
      tab_mayor.Delete
      tab_mayor.MoveNext
    Loop
    ultimo = -1
    For fila = 2 To grid_comi.Rows - 1
        ultimo = ultimo + 1
        tab_mayor.AddNew
        tab_mayor!TAB_CODCIA = PUB_CODCIA
        tab_mayor!TAB_TIPREG = PUB_TIPREG
        tab_mayor!TAB_NUMTAB = ultimo
        tab_mayor!tab_NOMLARGO = Trim(grid_comi.TextMatrix(fila, 3))
        tab_mayor!tab_nomcorto = Trim(grid_comi.TextMatrix(fila, 1))
        tab_mayor!TAB_CODART = val(grid_comi.TextMatrix(fila, 2))
        tab_mayor.Update
    Next fila
    CN.Execute "Commit Transaction", rdExecDirect
    On Error GoTo 0
    Screen.MousePointer = 0
    cmdAgregar.Enabled = False
    Exit Sub
End If


If PUB_TIPREG = 199 Then MsgBox "Solo hasta 4 tipos de bloqueos"
If PUB_TIPREG = 199 Then If Gridtablas.Rows >= 5 Then Exit Sub
wBAN = 1
If PUB_TIPREG = 2102 Then
 Mensaje = "Ingrese Codigo para el Turno :"
 Titulo = "Datos Necesario..."
 valorpred = " "
 wvalor = InputBox(Mensaje, Titulo, valorpred)
 If wvalor = "" Then
  wBAN = 0
  Exit Sub
 End If
 If val(wvalor) <= 0 Then
   MsgBox "Invalido Codigo de Turno .. Intente Nuevamente ..", 48, Pub_Titulo
   wBAN = 0
   Exit Sub
 End If
  WNUMTAB = wvalor
End If
Mensaje = "Ingrese la Descripción  :"
Titulo = "Datos Necesario..."
valorpred = " "
wvalor = InputBox(Mensaje, Titulo, valorpred)
If wvalor = "" Then
 wBAN = 0
 Exit Sub
End If
If Trim(wvalor) = "" Then
  MsgBox "Invalido Nombre .. Intente Nuevamente ..", 48, Pub_Titulo
  wBAN = 0
  Exit Sub
End If
PS_REP01(0) = PUB_CODCIA
PS_REP01(1) = PUB_TIPREG
PS_REP01(2) = wvalor
PS_REP01(3) = -1
'AGREGADO
If PUB_TIPREG = 123 Or PUB_TIPREG = 129 _
 Or PUB_TIPREG = 130 Or PUB_TIPREG = 131 Or PUB_TIPREG = 132 Or PUB_TIPREG = 20 Or PUB_TIPREG = 30 Then
  PS_REP01(4) = PUB_CODART
End If
llave_rep01.Requery
If Not llave_rep01.EOF Then
  MsgBox "Descripción Existe. Intente Nuevamente ..", 48, Pub_Titulo
  wBAN = 0
  Exit Sub
End If
cmdAgregar.Enabled = False
Command1.Enabled = False
WCODART = 0
'AGREGADO
If PUB_TIPREG = 123 Or PUB_TIPREG = 129 _
 Or PUB_TIPREG = 130 Or PUB_TIPREG = 131 Or PUB_TIPREG = 132 Or PUB_TIPREG = 20 Or PUB_TIPREG = 30 Then

 WCODART = PUB_CODART
End If
'PUB_CODCIA = LK_CODCIA
SQ_OPER = 2
LEER_TAB_LLAVE
tab_mayor.MoveLast
If tab_mayor.EOF Then
  ultimo = 0
Else
  ultimo = tab_mayor!TAB_NUMTAB
End If
ultimo = ultimo + 1
On Error GoTo SALE
If PUB_TIPREG = 50 Then
  wBAN = 1
  Mensaje = "Ingrese su Codigo de la SUNAT   :"
  Titulo = "Datos Necesario..."
  valorpred = " "
  WVALOR2 = InputBox(Mensaje, Titulo, valorpred)
  If WVALOR2 = "" Then
   wBAN = 0
   Exit Sub
  End If
  If Trim(WVALOR2) = "" Then
    MsgBox "Codigo Invalido .. Intente Nuevamente ..", 48, Pub_Titulo
    wBAN = 0
    Exit Sub
  End If
  ultimo = val(WVALOR2)
End If
If PUB_TIPREG = 333 Then
  wBAN = 1
  Mensaje = "Ingrese Codigo de Tipo de Negocio  :"
  Titulo = "Datos Necesario..."
  valorpred = " "
  WVALOR2 = InputBox(Mensaje, Titulo, valorpred)
  If WVALOR2 = "" Then
   wBAN = 0
   Exit Sub
  End If
  If Trim(WVALOR2) = "" Then
    MsgBox "Codigo Invalido .. Intente Nuevamente ..", 48, Pub_Titulo
    wBAN = 0
    Exit Sub
  End If
  If val(WVALOR2) = 0 Or val(WVALOR2) > 99 Then
     MsgBox "Codigo Invalido .. NO PERMITIDO..", 48, Pub_Titulo
     wBAN = 0
     Exit Sub
  End If
  ultimo = val(WVALOR2)
End If

   If Trim(Nulo_Valors(par_llave!par_art_cias)) <> "" And (PUB_TIPREG = 122 Or PUB_TIPREG = 123 Or PUB_TIPREG = 129 Or PUB_TIPREG = 130 Or PUB_TIPREG = 131 Or PUB_TIPREG = 132 Or PUB_TIPREG = 133) Then
        wcodciax = PUB_CODCIA
        PUB_CODCIA = LK_CODCIA
        xcuenta = 1
        For fila = 1 To 30
           PUB_CODCIA = Mid(Trim(par_llave!par_art_cias), xcuenta, 2)
           If Trim(PUB_CODCIA) = "" Then Exit For
           GoSub OTRO_CIA
           xcuenta = xcuenta + 2
        Next fila
        PUB_CODCIA = wcodciax
   Else
     GoSub OTRO_CIA
   End If
Item = val(FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 0))
Gridtablas.Rows = Gridtablas.Rows + 1
FrmDatArti.Gridtablas.RowHeight(Gridtablas.Rows - 1) = 280
If PUB_TIPREG = 2102 Then
FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 0) = WNUMTAB
Else
FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 0) = ultimo
End If
FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 1) = Left(wvalor, 40)
FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 2) = PUB_CODCIA
FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 3) = PUB_TIPREG
FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 4) = ultimo

cmdAgregar.Enabled = True
Command1.Enabled = True
wBAN = 0
Exit Sub

OTRO_CIA:
tab_mayor.AddNew
tab_mayor!TAB_CODCIA = PUB_CODCIA
tab_mayor!TAB_TIPREG = PUB_TIPREG
If PUB_TIPREG = 2102 Then
 tab_mayor!TAB_NUMTAB = WNUMTAB
Else
 tab_mayor!TAB_NUMTAB = ultimo
End If
 
tab_mayor!tab_NOMLARGO = Left(wvalor, 40)
tab_mayor!tab_nomcorto = Left(wvalor, 10)
tab_mayor!TAB_CODART = WCODART
tab_mayor.Update
Return



Exit Sub
ESCAPA:
    CN.Execute "Rollback Transaction", rdExecDirect
    MsgBox Err.Number & "  " & Err.Description & "  Intente Nuevamente."
    wBAN = 0
    cmdAgregar.Enabled = True
    Command1.Enabled = True

Exit Sub
SALE:
If Err.Number = 40002 Then
  MsgBox "Codigo Ya Existe.", 48, Pub_Titulo
Else
  MsgBox "Intente Nunevamente.", 48, Pub_Titulo
End If
cmdAgregar.Enabled = True
Command1.Enabled = True
wBAN = 0
Exit Sub
graba_cta:
lblmensa.Caption = "Actualizando . . ."
DoEvents
Screen.MousePointer = 11
wvalor = ""
llave_rep01.MoveFirst
For fila = 1 To Gridtablas.Rows - 1
   If Trim(Gridtablas.TextMatrix(fila, 3)) <> Trim(Gridtablas.TextMatrix(fila, 5)) Then
      llave_rep01.Edit
      llave_rep01!com_cuenta_cierre = Trim(Gridtablas.TextMatrix(fila, 3))
      Gridtablas.TextMatrix(fila, 4) = Trim(Gridtablas.TextMatrix(fila, 3))
      llave_rep01.Update
      wvalor = "A"
   End If
   llave_rep01.MoveNext
Next fila
Screen.MousePointer = 0
lblmensa.Caption = "[F2] ó [Enter] = Modificar  "
If wvalor = "A" Then
 MsgBox "Tabla de Cuentas de Cierre Actualizados.", 48, Pub_Titulo
End If
Exit Sub
If PUB_TIPREG = -10 Then GoTo Transportista
Transportista:

PS_REP01(0) = 0
llave_rep01.Requery
If llave_rep01.EOF Then
  Item = 0
Else
  llave_rep01.MoveLast
  Item = llave_rep01!TRN_KEY
End If
Item = Item + 1
If val(FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 0)) = Item Then
  MsgBox " Tiene un Transportista Agregado Modificar el Ultimo con el nombre : Nuevo Transportista.", 48, Pub_Titulo
  Exit Sub
End If
Gridtablas.Rows = Gridtablas.Rows + 1
FrmDatArti.Gridtablas.RowHeight(Gridtablas.Rows - 1) = 280
FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 0) = Item
FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 1) = "Nuevo Transportista "
FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 2) = " "
FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 3) = " "
FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 4) = " "
FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 5) = " "
Gridtablas.Row = Gridtablas.Rows - 1
Gridtablas.SetFocus
'********************
Graba_Direccion:
'********************
Dim strDir As String
Dim strRef As String
Dim fila2 As Integer
llave_rep01.Requery
strDir = Trim(InputBox("Ingrese la Dirección del Cliente", "Dirección"))

strRef = Trim(InputBox("Ingrese la Referencia del Cliente", "Referencia"))
If strDir = "" Then
 MsgBox "Dato ingresado no valido, Intentelo nuevamente", vbInformation, "Dirección"
 Exit Sub
End If
 CN.Execute "insert into dircli (codcia,codcli,cp,direc,ref) values('" & LK_CODCIA & "','" & val(frmCLI.Txt_key) & "','C','" & strDir & "','" & strRef & "')"
 
 'FrmDatArti.Gridtablas.AddItem "", Gridtablas.Rows
 fila2 = FrmDatArti.Gridtablas.Rows
 FrmDatArti.Gridtablas.Rows = FrmDatArti.Gridtablas.Rows + 1
 llave_rep01.Requery
 FrmDatArti.Gridtablas.TextMatrix(fila2 - 1, 0) = llave_rep01!DIRCLI
 FrmDatArti.Gridtablas.TextMatrix(fila2 - 1, 1) = strDir
 FrmDatArti.Gridtablas.TextMatrix(fila2 - 1, 2) = strRef
 Gridtablas.SetFocus
 
Exit Sub
End Sub

Private Sub Command1_Click()
If cmdAgregar.Enabled And Left(cmdAgregar.Caption, 2) = "&G" Then
   pub_mensaje = "Los Datos han cambiado desea Grabarlos ... ?"
   Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
   If Pub_Respuesta = vbNo Then
     GoTo fin
   End If
   cmdAgregar_Click
End If
fin:
Unload FrmDatArti
End Sub

Private Sub Form_Activate()
If PUB_TIPREG = -55 Then GoTo pasa_cuentas
If PUB_TIPREG = -10 Then GoTo Transportista
If PUB_TIPREG = -5 Then GoTo direcciones
If PUB_TIPREG = 444 Then
    cmdAgregar.Enabled = False
    FrmDatArti.Gridtablas.Clear
    FrmDatArti.Gridtablas.Visible = False
    lblmensa.Caption = "[F2] = Modificar        [Insert]= Agregar      [DEL] = Eliminar "
    cmdAgregar.Caption = "&Grabar"
    wBAN = 0
    FrmDatArti.grid_comi.Clear
    FrmDatArti.grid_comi.Cols = 7
    FrmDatArti.grid_comi.Rows = 3
    FrmDatArti.grid_comi.FixedRows = 2
    FrmDatArti.grid_comi.ColWidth(0) = 400
    FrmDatArti.grid_comi.ColWidth(1) = 800
    FrmDatArti.grid_comi.ColWidth(2) = 800
    FrmDatArti.grid_comi.ColWidth(3) = 1000
    FrmDatArti.grid_comi.ColWidth(4) = 1
    FrmDatArti.grid_comi.ColWidth(5) = 1
    FrmDatArti.grid_comi.ColWidth(6) = 1
    FrmDatArti.grid_comi.TextMatrix(0, 0) = "Item"
    FrmDatArti.grid_comi.TextMatrix(0, 1) = "Dia    "
    FrmDatArti.grid_comi.TextMatrix(1, 1) = "Inicial  "
    FrmDatArti.grid_comi.TextMatrix(0, 2) = "Dia    "
    FrmDatArti.grid_comi.TextMatrix(1, 2) = "Final    "
    FrmDatArti.grid_comi.TextMatrix(0, 3) = "(%)    "
    FrmDatArti.grid_comi.TextMatrix(1, 3) = "Comision "
    FrmDatArti.grid_comi.ColAlignment(1) = 7
    FrmDatArti.grid_comi.ColAlignment(2) = 7
    FrmDatArti.grid_comi.ColAlignment(3) = 7
    fila = 1
    SQ_OPER = 2
    If PUB_TIPREG = 123 Then
      SQ_OPER = 3
    End If
    LEER_TAB_LLAVE
    Do Until tab_mayor.EOF
     fila = fila + 1
     FrmDatArti.grid_comi.Rows = FrmDatArti.grid_comi.Rows + 1
     FrmDatArti.grid_comi.RowHeight(fila) = 280
     FrmDatArti.grid_comi.TextMatrix(fila, 0) = Nulo_Valor0(tab_mayor!TAB_NUMTAB) + 1
     FrmDatArti.grid_comi.TextMatrix(fila, 1) = Format(tab_mayor!tab_nomcorto, "0")
     FrmDatArti.grid_comi.TextMatrix(fila, 2) = Format(tab_mayor!TAB_CODART, "0")
     FrmDatArti.grid_comi.TextMatrix(fila, 3) = Format(tab_mayor!tab_NOMLARGO, "0.00")
     FrmDatArti.grid_comi.TextMatrix(fila, 4) = Nulo_Valor0(tab_mayor!TAB_CODCIA)
     FrmDatArti.grid_comi.TextMatrix(fila, 5) = Nulo_Valor0(tab_mayor!TAB_TIPREG)
     FrmDatArti.grid_comi.TextMatrix(fila, 6) = Nulo_Valor0(tab_mayor!TAB_NUMTAB)
     tab_mayor.MoveNext
    Loop
    FrmDatArti.grid_comi.Rows = FrmDatArti.grid_comi.Rows - 1
    FrmDatArti.grid_comi.Visible = True
    FrmDatArti.grid_comi.SetFocus
Else
    FrmDatArti.grid_comi.Clear
    FrmDatArti.grid_comi.Visible = False
    wBAN = 0
    FrmDatArti.Gridtablas.Clear
    FrmDatArti.Gridtablas.ColWidth(0) = 600
    FrmDatArti.Gridtablas.ColWidth(1) = 2500
    FrmDatArti.Gridtablas.ColWidth(2) = 1
    FrmDatArti.Gridtablas.ColWidth(3) = 1
    FrmDatArti.Gridtablas.ColWidth(4) = 1
    FrmDatArti.Gridtablas.TextMatrix(0, 0) = "Cod."
    FrmDatArti.Gridtablas.TextMatrix(0, 1) = "Descripcion"
    If PUB_TIPREG = 131 Then
     FrmDatArti.Gridtablas.Cols = FrmDatArti.Gridtablas.Cols + 1
'     FrmDatArti.Gridtablas.RowHeight(0) = 300
    End If
     
    If PUB_TIPREG = 35 And LK_EMP = "PAR" Then
       FrmDatArti.Gridtablas.Cols = FrmDatArti.Gridtablas.Cols + 1
       FrmDatArti.Gridtablas.ColWidth(5) = 600
       FrmDatArti.Gridtablas.TextMatrix(0, 5) = "Vend."
    End If
    fila = 0
   If PUB_TIPREG = 123 Or PUB_TIPREG = 129 Or PUB_TIPREG = 130 Or PUB_TIPREG = 131 Or PUB_TIPREG = 132 Or PUB_TIPREG = 20 Or PUB_TIPREG = 30 Then
     SQ_OPER = 3
     LEER_TAB_LLAVE
     FrmDatArti.Gridtablas.RowHeight(0) = 300
     FrmDatArti.Gridtablas.Rows = 1
     Do Until tab_menor.EOF
      fila = fila + 1
      FrmDatArti.Gridtablas.Rows = FrmDatArti.Gridtablas.Rows + 1
      FrmDatArti.Gridtablas.RowHeight(fila) = 280
      FrmDatArti.Gridtablas.TextMatrix(fila, 0) = Nulo_Valor0(tab_menor!TAB_NUMTAB)
      FrmDatArti.Gridtablas.TextMatrix(fila, 1) = Trim(Nulo_Valors(tab_menor!tab_NOMLARGO))
      FrmDatArti.Gridtablas.TextMatrix(fila, 2) = Nulo_Valors(tab_menor!TAB_CODCIA)
      FrmDatArti.Gridtablas.TextMatrix(fila, 3) = Nulo_Valor0(tab_menor!TAB_TIPREG)
      FrmDatArti.Gridtablas.TextMatrix(fila, 4) = Nulo_Valor0(tab_menor!TAB_NUMTAB)
      If PUB_TIPREG = 131 Then
        FrmDatArti.Gridtablas.TextMatrix(fila, 5) = Nulo_Valor0(tab_menor!TAB_CODART)
      End If
      If PUB_TIPREG = 35 And LK_EMP = "PAR" Then
        FrmDatArti.Gridtablas.TextMatrix(fila, 5) = Nulo_Valor0(tab_menor!TAB_CODART)
      End If
      If PUB_TIPREG = 222 Or PUB_TIPREG = 122 Then
        FrmDatArti.Gridtablas.TextMatrix(fila, 5) = val(Nulo_Valor0(tab_menor!tab_nomcorto))
        FrmDatArti.Gridtablas.TextMatrix(fila, 6) = val(Nulo_Valor0(tab_menor!TAB_contable2))
      End If
      
      tab_menor.MoveNext
     Loop
    Else
     If PUB_TIPREG = 222 Or PUB_TIPREG = 122 Then
        FrmDatArti.Gridtablas.Clear
        FrmDatArti.Gridtablas.Cols = 7
        FrmDatArti.Gridtablas.ColWidth(0) = 600
        FrmDatArti.Gridtablas.ColWidth(1) = 1800
        FrmDatArti.Gridtablas.ColWidth(2) = 1
        FrmDatArti.Gridtablas.ColWidth(3) = 1
        FrmDatArti.Gridtablas.ColWidth(4) = 1
        FrmDatArti.Gridtablas.ColWidth(5) = 1000
        FrmDatArti.Gridtablas.ColWidth(6) = 1
        If PUB_TIPREG = 122 Then FrmDatArti.Gridtablas.ColWidth(6) = 1000
        FrmDatArti.Gridtablas.TextMatrix(0, 0) = "Item"
        FrmDatArti.Gridtablas.TextMatrix(0, 1) = "Descripcion"
        FrmDatArti.Gridtablas.TextMatrix(0, 5) = "Cta.Contab.1"
        FrmDatArti.Gridtablas.TextMatrix(0, 6) = "Cta.Contab.2"
     End If
     SQ_OPER = 2
     LEER_TAB_LLAVE
     FrmDatArti.Gridtablas.RowHeight(0) = 300
     FrmDatArti.Gridtablas.Rows = 1
     Do Until tab_mayor.EOF
      fila = fila + 1
      FrmDatArti.Gridtablas.Rows = FrmDatArti.Gridtablas.Rows + 1
      FrmDatArti.Gridtablas.RowHeight(fila) = 280
      FrmDatArti.Gridtablas.TextMatrix(fila, 0) = Nulo_Valor0(tab_mayor!TAB_NUMTAB)
      FrmDatArti.Gridtablas.TextMatrix(fila, 1) = Trim(Nulo_Valors(tab_mayor!tab_NOMLARGO))
      FrmDatArti.Gridtablas.TextMatrix(fila, 2) = Nulo_Valors(tab_mayor!TAB_CODCIA)
      FrmDatArti.Gridtablas.TextMatrix(fila, 3) = Nulo_Valor0(tab_mayor!TAB_TIPREG)
      FrmDatArti.Gridtablas.TextMatrix(fila, 4) = Nulo_Valor0(tab_mayor!TAB_NUMTAB)
      If PUB_TIPREG = 131 Then
        FrmDatArti.Gridtablas.TextMatrix(fila, 5) = Nulo_Valor0(tab_mayor!TAB_CODART)
      End If
      If PUB_TIPREG = 35 And LK_EMP = "PAR" Then
        FrmDatArti.Gridtablas.TextMatrix(fila, 5) = Nulo_Valor0(tab_mayor!TAB_CODART)
      End If
      If PUB_TIPREG = 222 Or PUB_TIPREG = 122 Then
        FrmDatArti.Gridtablas.TextMatrix(fila, 5) = val(Nulo_Valor0(tab_mayor!tab_nomcorto))
        FrmDatArti.Gridtablas.TextMatrix(fila, 6) = val(Nulo_Valor0(tab_mayor!TAB_contable2))
      End If
      tab_mayor.MoveNext
     Loop
    End If
    FrmDatArti.Gridtablas.Visible = True
    FrmDatArti.Gridtablas.SetFocus
End If
Exit Sub
'************************
'agregado
'************************
direcciones:
  lblmensa.Caption = "[F2] ó [Enter] = Modificar  "
  cmdAgregar.Caption = "Gragar Dir."
  FrmDatArti.Gridtablas.Cols = 3
  FrmDatArti.Gridtablas.Clear
  FrmDatArti.Gridtablas.ColWidth(0) = 500
  FrmDatArti.Gridtablas.ColWidth(1) = 2000
  FrmDatArti.Gridtablas.ColWidth(2) = 2000
  FrmDatArti.Gridtablas.TextMatrix(0, 0) = "Item."
  FrmDatArti.Gridtablas.TextMatrix(0, 1) = "Direccion"
  FrmDatArti.Gridtablas.TextMatrix(0, 2) = "Referencia"
  
  PS_REP01(0) = LK_CODCIA
  PS_REP01(1) = val(frmCLI.Txt_key)
  PS_REP01(2) = "C"
  llave_rep01.Requery
  FrmDatArti.Gridtablas.Visible = True
  If llave_rep01.EOF Then Exit Sub
  fila = 0
  Gridtablas.Rows = 1
  Do Until llave_rep01.EOF
   fila = fila + 1
   Gridtablas.Rows = Gridtablas.Rows + 1
   FrmDatArti.Gridtablas.TextMatrix(fila, 0) = llave_rep01!DIRCLI
   FrmDatArti.Gridtablas.TextMatrix(fila, 1) = Trim(llave_rep01!direc)
   FrmDatArti.Gridtablas.TextMatrix(fila, 2) = Trim(llave_rep01!ref)
   llave_rep01.MoveNext
  Loop
  Gridtablas.Rows = Gridtablas.Rows + 1
  FrmDatArti.Gridtablas.SetFocus

Exit Sub
pasa_cuentas:
lblmensa.Caption = "[F2] ó [Enter] = Modificar  "
cmdAgregar.Caption = "Gragar Cta."
FrmDatArti.Width = FrmDatArti.Width + 1700
FrmDatArti.Gridtablas.Width = FrmDatArti.Gridtablas.Width + 1700
FrmDatArti.Gridtablas.Cols = 6
FrmDatArti.Gridtablas.Clear
FrmDatArti.Gridtablas.ColWidth(0) = 500
FrmDatArti.Gridtablas.ColWidth(1) = 900
FrmDatArti.Gridtablas.ColWidth(2) = 1500
FrmDatArti.Gridtablas.ColWidth(3) = 900
FrmDatArti.Gridtablas.ColWidth(4) = 1500
FrmDatArti.Gridtablas.ColWidth(5) = 1
FrmDatArti.Gridtablas.TextMatrix(0, 0) = "Item."
FrmDatArti.Gridtablas.TextMatrix(0, 1) = "Ctas.Contable."
FrmDatArti.Gridtablas.TextMatrix(0, 2) = "Descripción"
FrmDatArti.Gridtablas.TextMatrix(0, 3) = "Ctas.Cierre"
FrmDatArti.Gridtablas.TextMatrix(0, 4) = "Descripción"
pub_cadena = "SELECT COM_CUENTA,COM_DESCRIPCION, COM_CUENTA_CIERRE  FROM COMAEST WHERE COM_CODCIA = ? AND COM_NIVEL = ? and (COM_TIPO_CTA < 1 or  COM_TIPO_CTA > 5)  ORDER BY COM_CUENTA"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01.rdoParameters(0) = " "
PS_REP01.rdoParameters(1) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = cop_llave!cop_nivel_max
llave_rep01.Requery
If llave_rep01.EOF Then Exit Sub
fila = 0
Gridtablas.Rows = 1
Do Until llave_rep01.EOF
 fila = fila + 1
 Gridtablas.Rows = Gridtablas.Rows + 1
 FrmDatArti.Gridtablas.TextMatrix(fila, 0) = fila
 FrmDatArti.Gridtablas.TextMatrix(fila, 1) = Trim(llave_rep01!com_cuenta)
 FrmDatArti.Gridtablas.TextMatrix(fila, 2) = Trim(llave_rep01!com_descripcion)
 FrmDatArti.Gridtablas.TextMatrix(fila, 3) = Trim(Nulo_Valors(llave_rep01!com_cuenta_cierre))
 FrmDatArti.Gridtablas.TextMatrix(fila, 4) = "..."
 FrmDatArti.Gridtablas.TextMatrix(fila, 5) = Trim(FrmDatArti.Gridtablas.TextMatrix(fila, 3))
 llave_rep01.MoveNext
Loop
FrmDatArti.Gridtablas.Visible = True
FrmDatArti.Gridtablas.SetFocus

Exit Sub

Transportista:
pub_cadena = "SELECT * FROM TRANSPORTE WHERE TRN_KEY >= ? ORDER BY TRN_KEY"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = 0
cmdAgregar.Enabled = True
FrmDatArti.Gridtablas.Clear
FrmDatArti.Gridtablas.Visible = False
lblmensa.Caption = "[F2] = Modificar        [Insert]= Agregar      [DEL] = Eliminar "
wBAN = 0

FrmDatArti.Gridtablas.Clear
FrmDatArti.Gridtablas.Cols = 10
FrmDatArti.Gridtablas.Rows = 3
FrmDatArti.Gridtablas.FixedRows = 2
FrmDatArti.Gridtablas.ColWidth(0) = 400
FrmDatArti.Gridtablas.ColWidth(1) = 2500
FrmDatArti.Gridtablas.ColWidth(2) = 2500
FrmDatArti.Gridtablas.ColWidth(3) = 1000
FrmDatArti.Gridtablas.ColWidth(4) = 1000
FrmDatArti.Gridtablas.ColWidth(5) = 1000
FrmDatArti.Gridtablas.ColWidth(6) = 2500
FrmDatArti.Gridtablas.ColWidth(7) = 2500
FrmDatArti.Gridtablas.ColWidth(8) = 1000
FrmDatArti.Gridtablas.ColWidth(9) = 1000

FrmDatArti.Gridtablas.TextMatrix(0, 0) = "Item"
FrmDatArti.Gridtablas.TextMatrix(0, 1) = "Nombre /Razón social"
FrmDatArti.Gridtablas.TextMatrix(0, 2) = "Domicilio "
FrmDatArti.Gridtablas.TextMatrix(0, 3) = "R.U.C."
FrmDatArti.Gridtablas.TextMatrix(0, 4) = "D.N.I."
FrmDatArti.Gridtablas.TextMatrix(0, 5) = "Placa"
FrmDatArti.Gridtablas.TextMatrix(0, 6) = "Nombre Chofer"
FrmDatArti.Gridtablas.TextMatrix(0, 7) = "Dirección Chofer"
FrmDatArti.Gridtablas.TextMatrix(0, 8) = "Nro. Brevete"
FrmDatArti.Gridtablas.TextMatrix(0, 9) = "D.N.I. "

llave_rep01.Requery
If llave_rep01.EOF Then
 Gridtablas.Visible = True
 Gridtablas.Rows = Gridtablas.Rows - 1
 Exit Sub
End If
fila = 0
Gridtablas.Rows = 1
Do Until llave_rep01.EOF
 fila = fila + 1
 Gridtablas.Rows = Gridtablas.Rows + 1
 FrmDatArti.Gridtablas.RowHeight(fila) = 280
 FrmDatArti.Gridtablas.TextMatrix(fila, 0) = llave_rep01!TRN_KEY
 FrmDatArti.Gridtablas.TextMatrix(fila, 1) = Trim(llave_rep01!TRN_NOMBRE)
 FrmDatArti.Gridtablas.TextMatrix(fila, 2) = Trim(llave_rep01!TRN_DIRECCION)
 FrmDatArti.Gridtablas.TextMatrix(fila, 3) = Trim(Nulo_Valors(llave_rep01!TRN_RUC))
 FrmDatArti.Gridtablas.TextMatrix(fila, 4) = Trim(Nulo_Valors(llave_rep01!TRN_DNI))
 FrmDatArti.Gridtablas.TextMatrix(fila, 5) = Trim(Nulo_Valors(llave_rep01!TRN_PLACA))
 
 FrmDatArti.Gridtablas.TextMatrix(fila, 6) = Trim(Nulo_Valors(llave_rep01!TRN_CHOFER))
 FrmDatArti.Gridtablas.TextMatrix(fila, 7) = Trim(Nulo_Valors(llave_rep01!TRN_DIR_CHOFER))
 FrmDatArti.Gridtablas.TextMatrix(fila, 8) = Trim(Nulo_Valors(llave_rep01!TRN_BREVETE))
 FrmDatArti.Gridtablas.TextMatrix(fila, 9) = Trim(Nulo_Valors(llave_rep01!TRN_DNI_CHOFER))
 
 llave_rep01.MoveNext
Loop
FrmDatArti.Gridtablas.Visible = True
FrmDatArti.Gridtablas.SetFocus
FrmDatArti.Gridtablas.SetFocus

End Sub

Private Sub Form_Load()
Dim SQL As String

SQL = "SELECT * FROM DIRCLI WHERE CODCIA=? AND CODCLI=? AND CP=?"



If PUB_TIPREG = 123 Or PUB_TIPREG = 129 _
 Or PUB_TIPREG = 130 Or PUB_TIPREG = 131 Or PUB_TIPREG = 132 Or PUB_TIPREG = 20 Or PUB_TIPREG = 30 Then
  pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_TIPREG = ? AND TAB_NOMLARGO = ? AND TAB_NUMTAB <> ? AND TAB_CODART = ? ORDER BY TAB_CODCIA"
Else
  pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_TIPREG = ? AND TAB_NOMLARGO = ? AND TAB_NUMTAB <> ?  ORDER BY TAB_CODCIA"
End If
If PUB_TIPREG = -5 Then
 
 Set PS_REP01 = CN.CreateQuery("", SQL)
  PS_REP01.rdoParameters(0) = " "
  PS_REP01.rdoParameters(1) = 0
  PS_REP01.rdoParameters(2) = " "

Else
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
  PS_REP01.rdoParameters(0) = " "
  PS_REP01.rdoParameters(1) = 0
  PS_REP01.rdoParameters(2) = " "
  PS_REP01.rdoParameters(3) = 0
End If


If PUB_TIPREG = 123 Or PUB_TIPREG = 129 _
 Or PUB_TIPREG = 130 Or PUB_TIPREG = 131 Or PUB_TIPREG = 132 Or PUB_TIPREG = 20 Or PUB_TIPREG = 30 Then
 PS_REP01.rdoParameters(4) = 0
End If
'*******************************
'abrir el conjunto de registros
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

End Sub

Public Sub ElGrid_Click(wsGrid As MSFlexGrid, wsTexto As TextBox)
On Error GoTo SALE
wsTexto.Left = wsGrid.Left + wsGrid.CellLeft
wsTexto.Width = wsGrid.CellWidth
wsTexto.Top = wsGrid.Top + wsGrid.CellTop
wsTexto.Tag = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
wsTexto.Text = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
wsTexto.Visible = False
Exit Sub
SALE:
End Sub
Public Sub ElGrid_EnterCell(wsGrid As MSFlexGrid, wsTexto As TextBox, Optional Bloq1, Optional Bloq2, Optional Bloq3, Optional Bloq4, Optional Bloq5)
On Error GoTo SALE
wsGrid.CellFontBold = True
wsTexto.Left = wsGrid.Left + wsGrid.CellLeft
wsTexto.Width = wsGrid.CellWidth
wsTexto.Top = wsGrid.Top + wsGrid.CellTop
wsTexto.Tag = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
wsTexto.Text = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
Flag_Bloq = ""
If Not IsMissing(Bloq1) Then
  If wsGrid.COL = Bloq1 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq2) Then
  If wsGrid.COL = Bloq2 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq3) Then
  If wsGrid.COL = Bloq3 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq4) Then
  If wsGrid.COL = Bloq4 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq5) Then
  If wsGrid.COL = Bloq5 Then
    Flag_Bloq = "A"
  End If
End If
Exit Sub
SALE:
End Sub
Public Sub ElGrid_KeyDown(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyCode)
Flag_F2 = ""
If Flag_Bloq = "A" Then
  wsKeyCode = 0
  Exit Sub
End If

If wsKeyCode <> 113 Then
 Exit Sub
End If
If wsTexto.Visible = False Then
  Flag_F2 = "A"
  ElGrid_DblClick wsGrid, wsTexto
End If
End Sub
Public Sub ElGrid_KeyPress(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyAscii, Optional SaltaCol)
If wsKeyAscii = 27 Then
 Exit Sub
End If
If wsKeyAscii = 9 Or wsKeyAscii = 13 Then
  If Not IsMissing(SaltaCol) Then
    If wsGrid.COL = SaltaCol And wsGrid.Row <> wsGrid.Rows - 1 Then
       wsGrid.Row = wsGrid.Row + 1
       wsGrid.COL = wsGrid.FixedCols
       Exit Sub
    End If
  End If
  If wsGrid.COL <> wsGrid.Cols - 1 Then
    wsGrid.COL = wsGrid.COL + 1
  End If
  Exit Sub
End If
If Flag_Bloq = "A" Then
 wsKeyAscii = 0
 Exit Sub
End If

Dim cade
wsTexto.FontBold = True
wsTexto.ForeColor = QBColor(1)

wsTexto.Visible = True
If wsTexto.Enabled = True And wsTexto.Visible = True Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
End If
cade = Chr(wsKeyAscii)
SendKeys cade, True

End Sub
Private Sub ElGrid_LeaveCell(wsGrid As MSFlexGrid, wsTexto As TextBox)
If Flag_Consis = "A" Then
 wsTexto.FontBold = True
 wsTexto.ForeColor = QBColor(12)
 wsTexto.Visible = True
 If wsTexto.Enabled = True And wsTexto.Visible = True Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
 End If
 Exit Sub
End If
If Left(Trim(wsGrid.Text), 1) = "-" Then
 wsGrid.CellForeColor = QBColor(12)
Else
 wsGrid.CellForeColor = QBColor(0)
End If

wsGrid.CellFontBold = False
End Sub
Private Sub ElGrid_DblClick(wsGrid As MSFlexGrid, wsTexto As TextBox)
If Flag_Bloq = "A" Then
  Exit Sub
End If
wsTexto.FontBold = True
wsTexto.ForeColor = QBColor(12)
wsTexto.Visible = True
If wsTexto.Enabled = True And wsTexto.Visible = True Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
End If
End Sub
Private Sub ElGrid_GotFocus(wsGrid As MSFlexGrid, wsTexto As TextBox)
ElGrid_Click wsGrid, wsTexto
End Sub
Private Sub TEXTO_LosFocus(wsGrid As MSFlexGrid, wsTexto As TextBox)
ElGrid_Click wsGrid, wsTexto
End Sub
Public Sub TEXTO_KeyDown(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyCode As Integer, Optional SaltaCol)
If wsKeyCode = 46 Then
  TEXTO_KeyPress wsGrid, wsTexto, 46 '  wsKeyCode = 46
End If
If wsKeyCode = 40 Or wsKeyCode = 37 Or wsKeyCode = 39 Or wsKeyCode = 38 Then
 If Flag_F2 = "A" Then
   Exit Sub
 End If
 If Flag_Consis = "A" Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
   Beep
   Exit Sub
 End If
 If Gridtablas.Visible = False Then
   If grid_comi.COL <= 2 Then
    wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL) = Format(val(wsTexto.Text), "0")
   Else
    wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL) = Format(val(wsTexto.Text), "0.00")
   End If
   cmdAgregar.Enabled = True
 Else
   wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL) = wsTexto.Text
 End If
 wsTexto.Visible = False
 If wsKeyCode = 40 Then ' ABAJO
  If wsGrid.Row <> wsGrid.Rows - 1 Then
     wsGrid.Row = wsGrid.Row + 1
  End If
 End If
 If wsKeyCode = 38 Then ' arriba
  If wsGrid.Row <> wsGrid.FixedRows Then
     wsGrid.Row = wsGrid.Row - 1
  End If
 End If
 If wsKeyCode = 37 Then ' isquierda
  If wsGrid.COL <> wsGrid.FixedCols Then
   '  wsGrid.Col = wsGrid.Col - 1
  End If
 End If
 If wsKeyCode = 39 Then ' derecha
  If Not IsMissing(SaltaCol) Then
      If wsGrid.COL = SaltaCol And wsGrid.Row <> wsGrid.Rows - 1 Then
   '     wsGrid.Row = wsGrid.Row + 1
   '     wsGrid.Col = wsGrid.FixedCols
        GoTo wsfinal
      End If
  End If
  If wsGrid.COL <> wsGrid.Cols - 1 Then
  '   wsGrid.Col = wsGrid.Col + 1
  End If
 End If
wsfinal:
 wsTexto.FontBold = False
 wsTexto.ForeColor = QBColor(0)
 wsTexto.Text = ""
 wsGrid.SetFocus
End If

End Sub

Public Sub TEXTO_KeyPress(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyAscii As Integer, Optional SaltaCol, Optional ConsisCol1, Optional ConsisVal1, Optional ConsisCol2, Optional ConsisVal2, Optional ConsisCol3, Optional ConsisVal3, Optional ConsisCol4, Optional ConsisVal4, Optional ConsisCol5, Optional ConsisVal5, Optional ConsisCol6, Optional ConsisVal6, Optional ConsisCol7, Optional ConsisVal7, Optional ConsisCol8, Optional ConsisVal8, Optional ConsisCol9, Optional ConsisVal9, Optional ConsisCol10, Optional ConsisVal10)
If wsKeyAscii = 13 Or wsKeyAscii = 9 Then
  Flag_F2 = ""
  TEXTO_KeyDown wsGrid, wsTexto, 39, SaltaCol
  Exit Sub
End If
If wsKeyAscii = 27 Then
  ElGrid_Click wsGrid, wsTexto
  wsGrid.SetFocus
End If
If Not IsMissing(ConsisCol1) And Not IsMissing(ConsisVal1) Then
  If wsGrid.COL = ConsisCol1 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal1, ConsisCol1
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol2) And Not IsMissing(ConsisVal2) Then
  If wsGrid.COL = ConsisCol2 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal2, ConsisCol2
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol3) And Not IsMissing(ConsisVal3) Then
  If wsGrid.COL = ConsisCol3 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal3, ConsisCol3
   Exit Sub
  End If
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'PUB_TIPREG = -3

End Sub

Private Sub gridtablas_Click()
ElGrid_Click Gridtablas, TEXTOVAR
End Sub

Private Sub gridtablas_DblClick()
If PUB_TIPREG = -55 Then Exit Sub
ElGrid_DblClick Gridtablas, TEXTOVAR
End Sub

Private Sub gridtablas_EnterCell()
ElGrid_EnterCell Gridtablas, TEXTOVAR, 2, 3, 4
End Sub

Private Sub gridtablas_GotFocus()
ElGrid_GotFocus Gridtablas, TEXTOVAR
End Sub

Private Sub gridtablas_KeyDown(KeyCode As Integer, Shift As Integer)
If PUB_TIPREG = -55 Then
  If Gridtablas.COL <> 3 Then Exit Sub
  If KeyCode = 45 Then Exit Sub
End If
If KeyCode = 13 And Gridtablas.COL = 1 Then
 Exit Sub
End If
If Gridtablas.Row = 0 Then
 Exit Sub
End If
If KeyCode = 45 Then
 cmdAgregar_Click
 Exit Sub
End If
If KeyCode = 46 And PUB_TIPREG = -10 Then GoTo borra_transportista
'**************************************
If KeyCode = 46 And PUB_TIPREG = -5 Then GoTo Delete_Direccion
'**************************************
If KeyCode = 46 And PUB_TIPREG <> -55 Then
  ' ELIMINACION SOLO PARA EL SUPERVISOR O ADMIN
    If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Then
    Else
       Exit Sub
    End If
    pub_mensaje = " Eliminar    : " & Trim(Gridtablas.TextMatrix(Gridtablas.Row, 1)) & " ¿Desea Continuar... ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
      Gridtablas.SetFocus
      Exit Sub
    End If
    SQ_OPER = 1
    PUB_CODCIA = FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Row, 2)
    PUB_TIPREG = FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Row, 3)
    PUB_NUMTAB = FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Row, 4)
    LEER_TAB_LLAVE
    If tab_llave.EOF Then
      Exit Sub
    End If
    tab_llave.Delete
    If Gridtablas.Rows = 2 Then
     Gridtablas.Rows = Gridtablas.Rows - 1
    Else
    Gridtablas.RemoveItem Gridtablas.Row
    End If
    Gridtablas.SetFocus
   
End If
ElGrid_KeyDown Gridtablas, TEXTOVAR, KeyCode
Exit Sub

borra_transportista:
pub_mensaje = " Eliminar    : " & Trim(Gridtablas.TextMatrix(Gridtablas.Row, 1)) & " ¿Desea Continuar... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
  Gridtablas.SetFocus
  Exit Sub
End If

PS_REP01(0) = Gridtablas.TextMatrix(Gridtablas.Row, 0)
llave_rep01.Requery
If Not llave_rep01.EOF Then
  llave_rep01.Delete
End If
If Gridtablas.Rows = 2 Then
   Gridtablas.Rows = Gridtablas.Rows - 1
Else
   Gridtablas.RemoveItem Gridtablas.Row
End If
Gridtablas.SetFocus
Exit Sub
'PROCEDIMIENTO PARA ELIMINAR UNA DIRECCION
Delete_Direccion:
  Dim SQL As String
  With FrmDatArti.Gridtablas
  SQL = "DELETE FROM DIRCLI WHERE CODCIA='" & LK_CODCIA & "' " & _
        "AND DIRCLI='" & val(.TextMatrix(.Row, 0)) & "' AND " & _
        "CODCLI='" & val(frmCLI.Txt_key) & "' AND CP='C'"
  End With
  On Error GoTo ErrorDelete
  CN.Execute SQL
  Gridtablas.RemoveItem Gridtablas.Row
  Gridtablas.SetFocus
  Exit Sub
ErrorDelete:
 MsgBox Err.Description
End Sub

Private Sub gridtablas_KeyPress(KeyAscii As Integer)
If PUB_TIPREG = -55 And Gridtablas.COL = 3 Then
  ElGrid_KeyDown Gridtablas, TEXTOVAR, 113
 Exit Sub
End If

If KeyAscii = 13 And Gridtablas.COL = 1 Then
 Exit Sub
End If
If KeyAscii = 27 Then
 Command1_Click
 Exit Sub
End If
End Sub

Private Sub gridtablas_LeaveCell()
ElGrid_LeaveCell Gridtablas, TEXTOVAR
End Sub

Private Sub gridtablas_Scroll()
If Gridtablas.RowIsVisible(Gridtablas.Row) = False Then
   TEXTOVAR.Visible = False
   Exit Sub
Else
  TEXTOVAR.Visible = False
End If

TEXTOVAR.Left = Gridtablas.Left + Gridtablas.CellLeft ' gridtablas.CellLeft  'Frame4.Left + gridtablas.CellLeft
TEXTOVAR.Width = Gridtablas.CellWidth
TEXTOVAR.Top = Gridtablas.Top + Gridtablas.CellTop ' Frame4.Top + gridtablas.Top + gridtablas.CellTop
TEXTOVAR.DataField = Gridtablas.COL
TEXTOVAR.Tag = Gridtablas.Row
TEXTOVAR.Text = Gridtablas.TextMatrix(Gridtablas.Row, Gridtablas.COL)

End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
Dim xcuenta As Integer
Dim wcodciax As String * 2
Dim SQL As String
If KeyAscii = 13 Then
  '******************************
  'AGREGADO
  '******************************
  If PUB_TIPREG = -5 Then
   If FrmDatArti.Gridtablas.COL = 1 Then
    If Not Trim(TEXTOVAR) = "" Then
     SQL = "UPDATE DIRCLI SET DIREC='" & Trim(TEXTOVAR) & "' " & _
     "WHERE CODCIA='" & LK_CODCIA & "' AND CODCLI='" & val(frmCLI.Txt_key) & "' AND DIRCLI='" & val(FrmDatArti.Gridtablas.TextMatrix(FrmDatArti.Gridtablas.Row, 0)) & "'"
     CN.Execute SQL
     FrmDatArti.Gridtablas.TextMatrix(FrmDatArti.Gridtablas.Row, FrmDatArti.Gridtablas.COL) = Trim(TEXTOVAR)
     TEXTOVAR.Visible = False
     Exit Sub
    Else
     TEXTOVAR.Visible = False
     Exit Sub
    End If
   ElseIf FrmDatArti.Gridtablas.COL = 2 Then
    If Not Trim(TEXTOVAR) = "" Then
     SQL = "UPDATE DIRCLI SET REF='" & Trim(TEXTOVAR) & "' " & _
     "WHERE CODCIA='" & LK_CODCIA & "' AND CODCLI='" & val(frmCLI.Txt_key) & "' AND DIRCLI='" & val(FrmDatArti.Gridtablas.TextMatrix(FrmDatArti.Gridtablas.Row, 0)) & "'"
     CN.Execute SQL
     FrmDatArti.Gridtablas.TextMatrix(FrmDatArti.Gridtablas.Row, FrmDatArti.Gridtablas.COL) = Trim(TEXTOVAR)
     TEXTOVAR.Visible = False
     Exit Sub
    Else
     TEXTOVAR.Visible = False
     Exit Sub
    End If
   End If
  End If
  
  If PUB_TIPREG = -55 Then
    If FrmDatArti.Gridtablas.COL = 3 Then GoTo cuentas
    Exit Sub
  End If
  If PUB_TIPREG = -10 Then
    GoTo transportistas
    Exit Sub
  End If
  If PUB_TIPREG = 35 And LK_EMP = "PAR" And FrmDatArti.Gridtablas.COL = 5 Then
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      PUB_CODVEN = val(TEXTOVAR.Text)
      LEER_VEN_LLAVE
      If ven_llave.EOF Then
         MsgBox "Vendedor No Existe ... ", 48, Pub_Titulo
         Exit Sub
      End If
  End If



 PS_REP01(0) = PUB_CODCIA
 PS_REP01(1) = PUB_TIPREG
 PS_REP01(2) = TEXTOVAR.Text
 
 If val(FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Row, 4)) <> 0 Then
  PS_REP01(3) = val(FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Row, 4))
 Else
  PS_REP01(3) = -1
 End If
 llave_rep01.Requery
 If Not llave_rep01.EOF Then
   MsgBox "Descripción Existe. Intente Nuevamente ..", 48, Pub_Titulo
   Azul TEXTOVAR, TEXTOVAR
   Exit Sub
 End If
End If

If Gridtablas.Visible = False Then
 If KeyAscii = 13 And grid_comi.COL = 1 And val(TEXTOVAR.Text) > 365 Then
   Azul TEXTOVAR, TEXTOVAR
   Exit Sub
 End If
 If KeyAscii = 13 And grid_comi.COL = 2 And val(TEXTOVAR.Text) > 365 Then
   Azul TEXTOVAR, TEXTOVAR
   Exit Sub
 End If
 If KeyAscii = 13 And grid_comi.COL = 3 And val(TEXTOVAR.Text) > 100 Then
   Azul TEXTOVAR, TEXTOVAR
   Exit Sub
 End If
 TEXTO_KeyPress grid_comi, TEXTOVAR, KeyAscii, 3, 1, 1, 2, 1, 3, 2
 Exit Sub
End If
If PUB_TIPREG = -10 Then
TEXTO_KeyPress Gridtablas, TEXTOVAR, KeyAscii
Else
TEXTO_KeyPress Gridtablas, TEXTOVAR, KeyAscii, 2, 5, 2, 6, 2
End If
If KeyAscii = 13 Then
  If Gridtablas.Visible Then
    If Trim(Nulo_Valors(par_llave!par_art_cias)) <> "" And (PUB_TIPREG = 122 Or PUB_TIPREG = 123 Or PUB_TIPREG = 129 Or PUB_TIPREG = 130 Or PUB_TIPREG = 131 Or PUB_TIPREG = 132 Or PUB_TIPREG = 133) Then
       xcuenta = 1
       wcodciax = PUB_CODCIA
       For fila = 1 To 30
          PUB_CODCIA = Mid(Trim(par_llave!par_art_cias), xcuenta, 2)
          If Trim(PUB_CODCIA) = "" Then Exit For
          GoSub OTRA_CIA
          xcuenta = xcuenta + 2
       Next fila
       PUB_CODCIA = wcodciax
    Else
      PUB_CODCIA = FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Row, 2)
      GoSub OTRA_CIA
    End If

    ElseIf grid_comi.Visible Then
    
    End If
End If
Exit Sub

OTRA_CIA:
     SQ_OPER = 1
     PUB_TIPREG = FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Row, 3)
     PUB_NUMTAB = FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Row, 4)
     LEER_TAB_LLAVE
     If tab_llave.EOF Then
       tab_llave.AddNew
     Else
       tab_llave.Edit
     End If
     tab_llave!TAB_CODCIA = PUB_CODCIA
     tab_llave!TAB_TIPREG = PUB_TIPREG
     tab_llave!TAB_NUMTAB = PUB_NUMTAB
     tab_llave!tab_NOMLARGO = FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Row, 1)
     tab_llave!tab_nomcorto = FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Row, 1)
     If PUB_TIPREG = 131 Then tab_llave!TAB_CODART = val(FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Row, 5))
     
     If PUB_TIPREG = 222 Or PUB_TIPREG = 122 Then
      tab_llave!tab_nomcorto = FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Row, 5)
      tab_llave!TAB_contable2 = FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Row, 6)
     End If
     If PUB_TIPREG = 35 And LK_EMP = "PAR" Then
       tab_llave!TAB_CODART = val(FrmDatArti.Gridtablas.TextMatrix(Gridtablas.Row, 5))
     End If
     tab_llave.Update
Return
cuentas:
If Trim(TEXTOVAR.Text) = "" Then GoTo pasito
SQ_OPER = 1
PUB_CUENTA = Trim(TEXTOVAR.Text)
LEER_COM_LLAVE
If com_llave.EOF Then
 MsgBox "Cuenta No Existe .  Intente Nuevamente . . .", 48, Pub_Titulo
 Azul TEXTOVAR, TEXTOVAR
 Exit Sub
End If
If val(com_llave!COM_NIVEL) <> val(cop_llave!cop_nivel_max) Then
 MsgBox "Cuenta no Procede , No es ultimo nivel. . .", 48, Pub_Titulo
 Azul TEXTOVAR, TEXTOVAR
 Exit Sub
End If
If val(com_llave!COM_TIPO_CTA) >= 1 And val(com_llave!COM_TIPO_CTA) <= 5 Then
  MsgBox "Cuenta no Procede , es cuenta de Balance. . .", 48, Pub_Titulo
  Azul TEXTOVAR, TEXTOVAR
  Exit Sub
End If
Gridtablas.TextMatrix(Gridtablas.Row, 4) = com_llave!com_descripcion
pasito:
TEXTO_KeyPress Gridtablas, TEXTOVAR, KeyAscii, 2, 5, 2, 6, 2
If Gridtablas.Row <> Gridtablas.Rows - 1 Then
  Gridtablas.Row = Gridtablas.Row + 1
End If
Exit Sub

transportistas:
PS_REP01(0) = Gridtablas.TextMatrix(Gridtablas.Row, 0)
llave_rep01.Requery
If llave_rep01.EOF Then
  llave_rep01.AddNew
Else
  llave_rep01.Edit
End If
TEXTO_KeyPress Gridtablas, TEXTOVAR, KeyAscii, 2, 5, 2, 6, 2
llave_rep01!TRN_KEY = FrmDatArti.Gridtablas.TextMatrix(FrmDatArti.Gridtablas.Row, 0)
llave_rep01!TRN_NOMBRE = FrmDatArti.Gridtablas.TextMatrix(FrmDatArti.Gridtablas.Row, 1)
llave_rep01!TRN_DIRECCION = FrmDatArti.Gridtablas.TextMatrix(FrmDatArti.Gridtablas.Row, 2)
llave_rep01!TRN_RUC = FrmDatArti.Gridtablas.TextMatrix(FrmDatArti.Gridtablas.Row, 3)
llave_rep01!TRN_DNI = FrmDatArti.Gridtablas.TextMatrix(FrmDatArti.Gridtablas.Row, 4)
llave_rep01!TRN_PLACA = FrmDatArti.Gridtablas.TextMatrix(FrmDatArti.Gridtablas.Row, 5)

llave_rep01!TRN_CHOFER = FrmDatArti.Gridtablas.TextMatrix(FrmDatArti.Gridtablas.Row, 6)
llave_rep01!TRN_DIR_CHOFER = FrmDatArti.Gridtablas.TextMatrix(FrmDatArti.Gridtablas.Row, 7)
llave_rep01!TRN_BREVETE = FrmDatArti.Gridtablas.TextMatrix(FrmDatArti.Gridtablas.Row, 8)
llave_rep01!TRN_DNI_CHOFER = FrmDatArti.Gridtablas.TextMatrix(FrmDatArti.Gridtablas.Row, 9)
llave_rep01.Update

End Sub
Private Sub textovar_LostFocus()
If Gridtablas.Visible Then
 TEXTO_LosFocus Gridtablas, TEXTOVAR
Else
 TEXTO_LosFocus grid_comi, TEXTOVAR
End If
End Sub


' aqui enpi

Private Sub grid_comi_Click()
If grid_comi.Rows < 3 Then
 Exit Sub
End If
ElGrid_Click grid_comi, TEXTOVAR
End Sub

Private Sub grid_comi_DblClick()
If grid_comi.Rows < 3 Then
 Exit Sub
End If
ElGrid_DblClick grid_comi, TEXTOVAR
End Sub

Private Sub grid_comi_EnterCell()
If grid_comi.Rows < 3 Then
 Exit Sub
End If
ElGrid_EnterCell grid_comi, TEXTOVAR
End Sub

Private Sub grid_comi_GotFocus()
ElGrid_GotFocus grid_comi, TEXTOVAR
End Sub

Private Sub grid_comi_KeyDown(KeyCode As Integer, Shift As Integer)
If grid_comi.Rows < 3 And KeyCode = 113 Then
 Exit Sub
End If
If grid_comi.Rows < 3 And KeyCode = 46 Then
 Exit Sub
End If

If KeyCode = 45 Then
'Insetar nuevo
 If grid_comi.TextMatrix(FrmDatArti.grid_comi.Rows - 1, 1) = "0" And grid_comi.TextMatrix(FrmDatArti.grid_comi.Rows - 1, 2) = "0" And grid_comi.TextMatrix(FrmDatArti.grid_comi.Rows - 1, 3) = "0.00" Then
  MsgBox "No procede, se Agrego y esta en blanco. ", 48, Pub_Titulo
  Exit Sub
 End If
 FrmDatArti.grid_comi.Rows = FrmDatArti.grid_comi.Rows + 1
 FrmDatArti.grid_comi.RowHeight(FrmDatArti.grid_comi.Rows - 1) = 280
 FrmDatArti.grid_comi.TextMatrix(FrmDatArti.grid_comi.Rows - 1, 1) = "0"
 FrmDatArti.grid_comi.TextMatrix(FrmDatArti.grid_comi.Rows - 1, 2) = "0"
 FrmDatArti.grid_comi.TextMatrix(FrmDatArti.grid_comi.Rows - 1, 3) = "0.00"
 cmdAgregar.Enabled = True
 Exit Sub
End If
If KeyCode = 46 Then
   ' NO PROCEDE NINGUNA ELIMINACION
   Exit Sub
   pub_mensaje = " Eliminar    : " & Trim(grid_comi.TextMatrix(grid_comi.Row, 1)) & " ¿Desea Continuar... ?"
   Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
   If Pub_Respuesta = vbNo Then
      grid_comi.SetFocus
      Exit Sub
   End If
    If grid_comi.Rows = 3 Then
     grid_comi.Rows = grid_comi.Rows - 1
    Else
    grid_comi.RemoveItem grid_comi.Row
    End If
    cmdAgregar.Enabled = True
    grid_comi.SetFocus
End If

ElGrid_KeyDown grid_comi, TEXTOVAR, KeyCode

End Sub

Private Sub grid_comi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And grid_comi.COL = 1 Then
 Exit Sub
End If
If KeyAscii = 27 Then
 Command1_Click
 Exit Sub
End If
End Sub

Private Sub grid_comi_LeaveCell()
ElGrid_LeaveCell grid_comi, TEXTOVAR
End Sub

Private Sub grid_comi_Scroll()
If grid_comi.RowIsVisible(grid_comi.Row) = False Then
   TEXTOVAR.Visible = False
   Exit Sub
Else
  TEXTOVAR.Visible = False
End If

TEXTOVAR.Left = grid_comi.Left + grid_comi.CellLeft ' grid_comi.CellLeft  'Frame4.Left + grid_comi.CellLeft
TEXTOVAR.Width = grid_comi.CellWidth
TEXTOVAR.Top = grid_comi.Top + grid_comi.CellTop ' Frame4.Top + grid_comi.Top + grid_comi.CellTop
TEXTOVAR.DataField = grid_comi.COL
TEXTOVAR.Tag = grid_comi.Row
TEXTOVAR.Text = grid_comi.TextMatrix(grid_comi.Row, grid_comi.COL)

End Sub




Private Sub Consistencias(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyAscii As Integer, Optional ConsisVal, Optional ConsisCol)
  Static valor
  Dim car As String
  Flag_Consis = ""
  If ConsisVal = 2 Then ' NUMEROS CON DECIMALES
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If wsKeyAscii = 45 Then
      If wsTexto.Text <> "" Then
         Beep
         wsKeyAscii = 0
         Exit Sub
      End If
      Flag_Consis = "A"
    End If
    If wsKeyAscii = 46 Then
      If InStr(1, wsTexto.Text, ".") <> 0 Then
        Beep
        wsKeyAscii = 0
        Exit Sub
      End If
    End If
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And car <> "." And car <> "-" Then
          wsKeyAscii = 0
          Beep
          Exit Sub
        End If
    End If
  ElseIf ConsisVal = 1 Then ' NUMEROS ENTEROS
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And wsKeyAscii <> 32 And car <> "-" Then
          wsKeyAscii = 0
          Beep
        End If
      End If
  End If

End Sub

