VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmDatplac 
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4350
   Icon            =   "FrmDatplac.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox textovar 
      Height          =   285
      Left            =   3840
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grid_comi 
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6376
      _Version        =   327680
      Cols            =   5
   End
   Begin MSFlexGridLib.MSFlexGrid Gridtablas 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6376
      _Version        =   327680
      Cols            =   5
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Retornar"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdagregar 
      Caption         =   "&Agregar"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblmensa 
      Caption         =   "F2 = Editar                                           DELETE = Eliminar"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   4095
   End
End
Attribute VB_Name = "FrmDatplac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wBAN As Integer
Dim PSVE2_LLAVE  As rdoQuery
Dim clive2_llave As rdoResultset
Dim PSVE2_SECU As rdoQuery
Dim cliplac_ve2 As rdoResultset
Dim PSVE2_MAYOR As rdoQuery
Dim clive2_mayor As rdoResultset

Dim PSPLAC_LLAVE As rdoQuery
Dim cliplac_llave As rdoResultset
Dim PSPLAC_MAYOR As rdoQuery
Dim cliplac_mayor As rdoResultset
Dim PSPLAC_SECU As rdoQuery
Dim cliplac_secu   As rdoResultset

Dim PSART_ALT As rdoQuery
Dim PSART_KEY As rdoQuery
Dim artloc_alt As rdoResultset
Dim artloc_key As rdoResultset
Dim PSPRE_KEY As rdoQuery
Dim pre_key As rdoResultset



Private Sub cmdagregar_Click()
Dim WCODART As Integer
Dim wvalor
Dim ultimo As Integer
Dim Item As Integer
Dim Wflag
If textovar.Visible Then
  MsgBox "Esta Usted. Editanto ....!", 48, Pub_Titulo
  Exit Sub
End If
If PUB_TIPREG = 2301 Then
  GoTo AGREGA_DESCTO
End If

If CmdAgregar.Caption = "Gragar Cta." Then
  GoTo graba_cta
End If
If Left(CmdAgregar.Caption, 2) = "&G" Then
    Item = 0
    Wflag = ""
    For fila = 2 To grid_comi.Rows - 1
       If Wflag = "A" Then
          MsgBox "Verificar , El Ultimo rango debe ser parecido a unico ", 48, Pub_Titulo
          Exit Sub
       End If
       If Val(grid_comi.TextMatrix(fila, 1)) <> 0 And Val(grid_comi.TextMatrix(fila, 2)) = 0 Then
         Wflag = "A"
       ElseIf Val(grid_comi.TextMatrix(fila, 1)) > Val(grid_comi.TextMatrix(fila, 2)) Then
         MsgBox "Verificar , Dia Inicial no debe ser Mayor que la Final ", 48, Pub_Titulo
         Exit Sub
       End If
       If Not Val(grid_comi.TextMatrix(fila, 1)) >= Item Then
          MsgBox "Verificar , Dia Inicial  debe estar en Orden ascendente ", 48, Pub_Titulo
          Exit Sub
       End If
       Item = Val(grid_comi.TextMatrix(fila, 1))
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
        tab_mayor!tab_numtab = ultimo
        tab_mayor!tab_nomlargo = Trim(grid_comi.TextMatrix(fila, 3))
        tab_mayor!tab_nomcorto = Trim(grid_comi.TextMatrix(fila, 1))
        tab_mayor!tab_codART = Val(grid_comi.TextMatrix(fila, 2))
        tab_mayor.Update
    Next fila
    CN.Execute "Commit Transaction", rdExecDirect
    On Error GoTo 0
    Screen.MousePointer = 0
    CmdAgregar.Enabled = False
    Exit Sub
End If
If PUB_TIPREG = 199 Then MsgBox "Solo hasta 4 tipos de bloqueos"
If PUB_TIPREG = 199 Then If Gridtablas.Rows >= 5 Then Exit Sub
wBAN = 1
Mensaje = "Ingrese la Descripción  :"
titulo = "Datos Necesario..."
valorpred = " "
wvalor = InputBox(Mensaje, titulo, valorpred)
If wvalor = "" Then
 wBAN = 0
 Exit Sub
End If
If Trim(wvalor) = "" Then
  MsgBox "Invalido Nombre .. Intente Nuevamente ..", 48, Pub_Titulo
  wBAN = 0
  Exit Sub
End If
If PUB_TIPREG = -33 Then
  GoTo VENDEDOR
End If
If PUB_TIPREG = 2103 Then GoTo salta_ot
PSPLAC_LLAVE(0) = PUB_CODCIA
PSPLAC_LLAVE(1) = PUB_TIPREG
PSPLAC_LLAVE(2) = pu_codclie
PSPLAC_LLAVE(3) = wvalor ' FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Row, 4)
cliplac_llave.Requery
If Not cliplac_llave.EOF Then
    MsgBox "Descripción Existe. Intente Nuevamente ..", 48, Pub_Titulo
    If Gridtablas.Visible Then Gridtablas.SetFocus
    Exit Sub
End If
CmdAgregar.Enabled = False
Command1.Enabled = False
PSPLAC_SECU(0) = LK_CODCIA
PSPLAC_SECU(1) = PUB_TIPREG
cliplac_secu.Requery
cliplac_mayor.Requery
If cliplac_secu.EOF Then
  ultimo = 0
Else
  ultimo = cliplac_secu!tab_numtab
End If
ultimo = ultimo + 1
On Error GoTo SALE
PUB_TIPREG = 2101
cliplac_mayor.AddNew
cliplac_mayor!TAB_CODCIA = LK_CODCIA
cliplac_mayor!TAB_TIPREG = PUB_TIPREG
cliplac_mayor!tab_numtab = ultimo
cliplac_mayor!tab_nomlargo = Left(wvalor, 40)
cliplac_mayor!tab_nomcorto = Left(wvalor, 10)
cliplac_mayor!tab_codclie = pu_codclie
cliplac_mayor.Update
Item = Val(FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 0))
Gridtablas.Rows = Gridtablas.Rows + 1
FrmDatplac.Gridtablas.RowHeight(Gridtablas.Rows - 1) = 280
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 0) = ultimo
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 1) = Left(wvalor, 40)
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 2) = LK_CODCIA
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 3) = PUB_TIPREG
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 4) = ultimo
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 5) = pu_codclie

CmdAgregar.Enabled = True
Command1.Enabled = True
wBAN = 0
Exit Sub

VENDEDOR:
PSVE2_SECU(0) = LK_CODCIA
cliplac_ve2.Requery
If cliplac_ve2.EOF Then
  ultimo = 0
Else
  ultimo = cliplac_ve2!VE2_CODVEN
End If
 ultimo = ultimo + 1
 cliplac_ve2.AddNew
 cliplac_ve2!VE2_CODCIA = LK_CODCIA
 cliplac_ve2!VE2_CODVEN = ultimo
 cliplac_ve2!VE2_NOMBRE = wvalor
 cliplac_ve2.Update
Item = ultimo
Gridtablas.Rows = Gridtablas.Rows + 1
FrmDatplac.Gridtablas.RowHeight(Gridtablas.Rows - 1) = 280
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 0) = ultimo
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 1) = Left(wvalor, 40)
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 2) = LK_CODCIA
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 3) = ""

CmdAgregar.Enabled = True
Command1.Enabled = True
salta_ot:
SQ_OPER = 2
LEER_TAB_LLAVE
tab_mayor.MoveLast
If tab_mayor.EOF Then
  ultimo = 0
Else
  ultimo = tab_mayor!tab_numtab
End If
ultimo = ultimo + 1
tab_mayor.AddNew
tab_mayor!TAB_CODCIA = PUB_CODCIA
tab_mayor!TAB_TIPREG = PUB_TIPREG
tab_mayor!tab_numtab = ultimo
tab_mayor!tab_nomlargo = Left(wvalor, 40)
tab_mayor!tab_nomcorto = Left(wvalor, 10)
tab_mayor!tab_codART = 0
tab_mayor.Update
Item = Val(FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 0))
Gridtablas.Rows = Gridtablas.Rows + 1
FrmDatplac.Gridtablas.RowHeight(Gridtablas.Rows - 1) = 280
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 0) = ultimo
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 1) = Left(wvalor, 40)
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 2) = PUB_CODCIA
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 3) = PUB_TIPREG
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 4) = ultimo



CmdAgregar.Enabled = True
Command1.Enabled = True

Exit Sub

AGREGA_DESCTO:
wBAN = 1
Mensaje = "Ingrese la Codigo de Articulo :"
titulo = "Datos Necesario..."
valorpred = " "
wvalor = InputBox(Mensaje, titulo, valorpred)
If wvalor = "" Then
 wBAN = 0
 Exit Sub
End If
If Trim(wvalor) = "" Then
  MsgBox "Codigo No Procede.   Intente Nuevamente ..", 48, Pub_Titulo
  wBAN = 0
  Exit Sub
End If
PSART_ALT(0) = LK_CODCIA
PSART_ALT(1) = wvalor
artloc_alt.Requery
If artloc_alt.EOF Then
  MsgBox "Articulo No Existe ", 48, Pub_Titulo
  Exit Sub
End If
PSPRE_KEY(0) = LK_CODCIA
PSPRE_KEY(1) = artloc_alt!ART_KEY
If BUSCA(artloc_alt!ART_KEY) Then
  MsgBox "Codigo de Producto Existe en Lista.", 48, Pub_Titulo
  Exit Sub
End If
pre_key.Requery

CmdAgregar.Enabled = False
Command1.Enabled = False

PSVE2_SECU(0) = LK_CODCIA
cliplac_ve2.Requery
clive2_mayor.Requery
If cliplac_ve2.EOF Then
  ultimo = 0
Else
  ultimo = cliplac_ve2!tab_numtab
End If
ultimo = ultimo + 1


On Error GoTo SALE
PUB_TIPREG = 2301
clive2_mayor.AddNew
clive2_mayor!TAB_CODCIA = LK_CODCIA
clive2_mayor!TAB_TIPREG = PUB_TIPREG
clive2_mayor!tab_numtab = ultimo
clive2_mayor!tab_nomlargo = pre_key!PRE_PRE1
clive2_mayor!tab_nomcorto = pre_key!PRE_PRE1
clive2_mayor!tab_codclie = pu_codclie
clive2_mayor!tab_codART = artloc_alt!ART_KEY
clive2_mayor.Update
Item = Val(FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 0))
Gridtablas.Rows = Gridtablas.Rows + 1
FrmDatplac.Gridtablas.RowHeight(Gridtablas.Rows - 1) = 280
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 0) = ultimo
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 1) = artloc_alt!art_nombre
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 2) = pre_key!PRE_PRE1
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 3) = pre_key!PRE_PRE1
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 4) = pu_codclie
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 5) = ultimo
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 6) = PUB_TIPREG
FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 7) = artloc_alt!ART_KEY

'FrmDatplac.Gridtablas.TextMatrix(fila, 0) = ultimo ' Nulo_Valor0(clive2_mayor!TAB_NUMTAB)
'FrmDatplac.Gridtablas.TextMatrix(fila, 1) = artloc_alt!ART_nombre
'FrmDatplac.Gridtablas.TextMatrix(fila, 2) = "0.00"
'FrmDatplac.Gridtablas.TextMatrix(fila, 3) = Trim(Nulo_Valor0(clive2_mayor!TAB_NOMCORTO))
'FrmDatplac.Gridtablas.TextMatrix(fila, 4) = Trim(Nulo_Valor0(clive2_mayor!TAB_NUMTAB))
'FrmDatplac.Gridtablas.TextMatrix(fila, 5) = Trim(Nulo_Valor0(clive2_mayor!tab_tipreg))
'FrmDatplac.Gridtablas.TextMatrix(fila, 6) = Nulo_Valors(clive2_mayor!tab_codart)

CmdAgregar.Enabled = True
Command1.Enabled = True
wBAN = 0


Exit Sub



Exit Sub
ESCAPA:
    CN.Execute "Rollback Transaction", rdExecDirect
    MsgBox Err.Number & "  " & Err.Description & "  Intente Nuevamente."
    wBAN = 0
    CmdAgregar.Enabled = True
    Command1.Enabled = True

Exit Sub
SALE:
If Err.Number = 40002 Then
  MsgBox "Codigo Ya Existe.", 48, Pub_Titulo
Else
  MsgBox "Intente Nunevamente.", 48, Pub_Titulo
End If
MsgBox Err.Description, 48, Pub_Titulo
CmdAgregar.Enabled = True
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
End Sub

Private Sub Command1_Click()
If CmdAgregar.Enabled And Left(CmdAgregar.Caption, 2) = "&G" Then
   pub_mensaje = "Los Datos han cambiado desea Grabarlos ... ?"
   Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
   If Pub_Respuesta = vbNo Then
     GoTo FIN
   End If
   cmdagregar_Click
End If
FIN:
Unload FrmDatplac
End Sub

Private Sub Form_Activate()
If PUB_TIPREG = -33 Then
   GoTo pasa_vendedor
   Exit Sub
End If
If PUB_TIPREG = 2301 Then
   GoTo pasa_descto
   Exit Sub
End If
If PUB_TIPREG = 2103 Then
   GoTo pasa_otros
   Exit Sub
End If

    pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_TIPREG = ? AND TAB_CODCLIE = ? and TAB_NOMLARGO = ?  ORDER BY TAB_NOMLARGO"
    Set PSPLAC_LLAVE = CN.CreateQuery("", pub_cadena)
    PSPLAC_LLAVE.rdoParameters(0) = 0
    PSPLAC_LLAVE.rdoParameters(1) = 0
    PSPLAC_LLAVE.rdoParameters(2) = 0
    PSPLAC_LLAVE.rdoParameters(3) = 0
    Set cliplac_llave = PSPLAC_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
    
    pub_cadena = "SELECT TAB_NUMTAB FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_TIPREG = ? ORDER BY TAB_NUMTAB DESC"
    Set PSPLAC_SECU = CN.CreateQuery("", pub_cadena)
    PSPLAC_SECU.rdoParameters(0) = " "
    PSPLAC_SECU.rdoParameters(1) = 0
    PSPLAC_SECU.MaxRows = 1
    Set cliplac_secu = PSPLAC_SECU.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)
    
    pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_TIPREG = ? AND TAB_CODCLIE = ? ORDER BY TAB_NUMTAB"
    Set PSPLAC_MAYOR = CN.CreateQuery("", pub_cadena)
    PSPLAC_MAYOR.rdoParameters(0) = " "
    PSPLAC_MAYOR.rdoParameters(1) = 0
    PSPLAC_MAYOR.rdoParameters(2) = 0
    Set cliplac_mayor = PSPLAC_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
    
    FrmDatplac.grid_comi.Clear
    FrmDatplac.grid_comi.Visible = False
    wBAN = 0
    FrmDatplac.Gridtablas.Clear
    FrmDatplac.Gridtablas.Cols = 6
    FrmDatplac.Gridtablas.ColWidth(0) = 600
    FrmDatplac.Gridtablas.ColWidth(1) = 2500
    FrmDatplac.Gridtablas.ColWidth(2) = 1
    FrmDatplac.Gridtablas.ColWidth(3) = 1
    FrmDatplac.Gridtablas.ColWidth(4) = 1
    FrmDatplac.Gridtablas.ColWidth(5) = 1
    FrmDatplac.Gridtablas.TextMatrix(0, 0) = "Cod."
    FrmDatplac.Gridtablas.TextMatrix(0, 1) = "Descripcion"
    fila = 0
     PSPLAC_MAYOR(0) = LK_CODCIA
     PSPLAC_MAYOR(1) = 2101
     PSPLAC_MAYOR(2) = pu_codclie
     cliplac_mayor.Requery
     FrmDatplac.Gridtablas.RowHeight(0) = 300
     FrmDatplac.Gridtablas.Rows = 1
     Do Until cliplac_mayor.EOF
       fila = fila + 1
       FrmDatplac.Gridtablas.Rows = FrmDatplac.Gridtablas.Rows + 1
       FrmDatplac.Gridtablas.RowHeight(fila) = 280
       FrmDatplac.Gridtablas.TextMatrix(fila, 0) = Nulo_Valor0(cliplac_mayor!tab_numtab)
       FrmDatplac.Gridtablas.TextMatrix(fila, 1) = Trim(Nulo_Valors(cliplac_mayor!tab_nomlargo))
       FrmDatplac.Gridtablas.TextMatrix(fila, 2) = Nulo_Valors(cliplac_mayor!TAB_CODCIA)
       FrmDatplac.Gridtablas.TextMatrix(fila, 3) = Nulo_Valor0(cliplac_mayor!TAB_TIPREG)
       FrmDatplac.Gridtablas.TextMatrix(fila, 4) = Nulo_Valor0(cliplac_mayor!tab_numtab)
       FrmDatplac.Gridtablas.TextMatrix(fila, 5) = Nulo_Valor0(cliplac_mayor!tab_codART)
       cliplac_mayor.MoveNext
     Loop
     FrmDatplac.Gridtablas.Visible = True
     FrmDatplac.Gridtablas.SetFocus
     
Exit Sub
pasa_vendedor:
    pub_cadena = "SELECT * FROM VEMAEST2 WHERE VE2_CODCIA = ? AND VE2_CODVEN = ? ORDER BY VE2_CODVEN"
    Set PSVE2_LLAVE = CN.CreateQuery("", pub_cadena)
    PSVE2_LLAVE.rdoParameters(0) = " "
    PSVE2_LLAVE.rdoParameters(1) = 0
    Set clive2_llave = PSVE2_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
    
    pub_cadena = "SELECT * FROM VEMAEST2 WHERE VE2_CODCIA = ? AND VE2_CODVEN > 0 ORDER BY VE2_CODVEN DESC"
    Set PSVE2_SECU = CN.CreateQuery("", pub_cadena)
    PSVE2_SECU.rdoParameters(0) = " "
    PSVE2_SECU.MaxRows = 1
    Set cliplac_ve2 = PSVE2_SECU.OpenResultset(rdOpenForwardOnly, rdConcurValues)
    
    pub_cadena = "SELECT * FROM VEMAEST2 WHERE VE2_CODCIA = ? AND VE2_CODVEN > ? ORDER BY VE2_CODVEN "
    Set PSVE2_MAYOR = CN.CreateQuery("", pub_cadena)
    PSVE2_MAYOR.rdoParameters(0) = " "
    PSVE2_MAYOR.rdoParameters(1) = 0
    Set clive2_mayor = PSVE2_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
 

     FrmDatplac.grid_comi.Clear
     FrmDatplac.grid_comi.Visible = False
     wBAN = 0
     FrmDatplac.Gridtablas.Clear
     FrmDatplac.Gridtablas.Cols = 4
     FrmDatplac.Gridtablas.ColWidth(0) = 600
     FrmDatplac.Gridtablas.ColWidth(1) = 2500
     FrmDatplac.Gridtablas.ColWidth(2) = 1
     FrmDatplac.Gridtablas.ColWidth(3) = 1

     FrmDatplac.Gridtablas.TextMatrix(0, 0) = "Cod."
     FrmDatplac.Gridtablas.TextMatrix(0, 1) = "Nombre"
     fila = 0
     PSVE2_MAYOR(0) = LK_CODCIA
     PSVE2_MAYOR(1) = 0
     clive2_mayor.Requery
     FrmDatplac.Gridtablas.RowHeight(0) = 300
     FrmDatplac.Gridtablas.Rows = 1
     Do Until clive2_mayor.EOF
       fila = fila + 1
       FrmDatplac.Gridtablas.Rows = FrmDatplac.Gridtablas.Rows + 1
       FrmDatplac.Gridtablas.RowHeight(fila) = 280
       FrmDatplac.Gridtablas.TextMatrix(fila, 0) = Nulo_Valor0(clive2_mayor!VE2_CODVEN)
       FrmDatplac.Gridtablas.TextMatrix(fila, 1) = Trim(Nulo_Valors(clive2_mayor!VE2_NOMBRE))
       FrmDatplac.Gridtablas.TextMatrix(fila, 2) = Nulo_Valors(clive2_mayor!VE2_CODCIA)
       clive2_mayor.MoveNext
     Loop
     FrmDatplac.Gridtablas.Visible = True
     FrmDatplac.Gridtablas.SetFocus
     
Exit Sub

pasa_descto:
    
    pub_cadena = "SELECT PRE_PRE1 FROM PRECIOS WHERE PRE_CODCIA = ? AND PRE_CODART = ? "
    Set PSPRE_KEY = CN.CreateQuery("", pub_cadena)
    PSPRE_KEY.rdoParameters(0) = " "
    PSPRE_KEY.rdoParameters(1) = 0
    Set pre_key = PSPRE_KEY.OpenResultset(rdOpenKeyset, rdConcurValues)
    
    pub_cadena = "SELECT ART_NOMBRE, ART_KEY FROM ARTI WHERE ART_CODCIA = ? AND ART_KEY = ? "
    Set PSART_KEY = CN.CreateQuery("", pub_cadena)
    PSART_KEY.rdoParameters(0) = " "
    PSART_KEY.rdoParameters(1) = 0
    Set artloc_key = PSART_KEY.OpenResultset(rdOpenKeyset, rdConcurValues)

    pub_cadena = "SELECT ART_NOMBRE, ART_KEY FROM ARTI WHERE ART_CODCIA = ? AND ART_ALTERNO = ? "
    Set PSART_ALT = CN.CreateQuery("", pub_cadena)
    PSART_ALT.rdoParameters(0) = " "
    PSART_ALT.rdoParameters(1) = " "
    Set artloc_alt = PSART_ALT.OpenResultset(rdOpenKeyset, rdConcurValues)

    pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_CODCLIE = ? AND TAB_CODART = ? AND TAB_TIPREG = 2301  ORDER BY TAB_NUMTAB"
    Set PSVE2_LLAVE = CN.CreateQuery("", pub_cadena)
    PSVE2_LLAVE.rdoParameters(0) = 0
    PSVE2_LLAVE.rdoParameters(1) = 0
    PSVE2_LLAVE.rdoParameters(2) = 0
    Set clive2_llave = PSVE2_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
    
    pub_cadena = "SELECT TAB_NUMTAB FROM TABLAS WHERE TAB_TIPREG = 2301 AND TAB_CODCIA = ? AND TAB_NUMTAB > 0 ORDER BY TAB_NUMTAB DESC"
    Set PSVE2_SECU = CN.CreateQuery("", pub_cadena)
    PSVE2_SECU.rdoParameters(0) = " "
    PSVE2_SECU.MaxRows = 1
    Set cliplac_ve2 = PSVE2_SECU.OpenResultset(rdOpenForwardOnly, rdConcurValues)
    
    pub_cadena = "SELECT * FROM TABLAS WHERE TAB_TIPREG = 2301 AND TAB_CODCIA = ? AND TAB_CODCLIE = ?  ORDER BY TAB_NUMTAB"
    Set PSVE2_MAYOR = CN.CreateQuery("", pub_cadena)
    PSVE2_MAYOR.rdoParameters(0) = 0
    PSVE2_MAYOR.rdoParameters(1) = 0
    Set clive2_mayor = PSVE2_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
 
     FrmDatplac.grid_comi.Clear
     FrmDatplac.grid_comi.Visible = False
     wBAN = 0
     FrmDatplac.Gridtablas.Clear
     FrmDatplac.Gridtablas.Cols = 8
     FrmDatplac.Gridtablas.ColWidth(0) = 400
     FrmDatplac.Gridtablas.ColWidth(1) = 1800
     FrmDatplac.Gridtablas.ColWidth(2) = 800
     FrmDatplac.Gridtablas.ColWidth(3) = 800
     FrmDatplac.Gridtablas.ColWidth(4) = 1
     FrmDatplac.Gridtablas.ColWidth(5) = 1
     FrmDatplac.Gridtablas.ColWidth(6) = 1
     FrmDatplac.Gridtablas.ColWidth(7) = 1

     FrmDatplac.Gridtablas.TextMatrix(0, 0) = "Item"
     FrmDatplac.Gridtablas.TextMatrix(0, 1) = "Producto"
     FrmDatplac.Gridtablas.TextMatrix(0, 2) = "P.Contado"
     FrmDatplac.Gridtablas.TextMatrix(0, 3) = "P.Credito"
     fila = 0
     PSVE2_MAYOR(0) = LK_CODCIA
     PSVE2_MAYOR(1) = pu_codclie
     clive2_mayor.Requery
     FrmDatplac.Gridtablas.RowHeight(0) = 300
     FrmDatplac.Gridtablas.Rows = 1
     Do Until clive2_mayor.EOF
       fila = fila + 1
       FrmDatplac.Gridtablas.Rows = FrmDatplac.Gridtablas.Rows + 1
       FrmDatplac.Gridtablas.RowHeight(fila) = 280
       FrmDatplac.Gridtablas.TextMatrix(fila, 0) = Nulo_Valor0(clive2_mayor!tab_numtab)
       PSART_KEY(0) = LK_CODCIA
       PSART_KEY(1) = Nulo_Valors(clive2_mayor!tab_codART)
       artloc_key.Requery
       If artloc_key.EOF Then
          clive2_mayor.Delete
          GoTo SIGUE
       End If
       FrmDatplac.Gridtablas.TextMatrix(fila, 1) = artloc_key!art_nombre
       FrmDatplac.Gridtablas.TextMatrix(fila, 2) = Format(Nulo_Valor0(clive2_mayor!tab_nomlargo), "0.00")
       FrmDatplac.Gridtablas.TextMatrix(fila, 3) = Format(Nulo_Valor0(clive2_mayor!tab_nomcorto), "0.00")
       FrmDatplac.Gridtablas.TextMatrix(fila, 4) = Trim(Nulo_Valor0(clive2_mayor!tab_codclie))
       FrmDatplac.Gridtablas.TextMatrix(fila, 5) = Trim(Nulo_Valor0(clive2_mayor!tab_numtab))
       FrmDatplac.Gridtablas.TextMatrix(fila, 6) = Trim(Nulo_Valor0(clive2_mayor!TAB_TIPREG))
       FrmDatplac.Gridtablas.TextMatrix(fila, 7) = Trim(Nulo_Valor0(clive2_mayor!tab_codART))
SIGUE:
       clive2_mayor.MoveNext
     Loop
     FrmDatplac.Gridtablas.Visible = True
     FrmDatplac.Gridtablas.SetFocus
     
Exit Sub
pasa_otros:
 FrmDatplac.Gridtablas.Clear
     FrmDatplac.Gridtablas.Cols = 5
     FrmDatplac.Gridtablas.ColWidth(0) = 600
     FrmDatplac.Gridtablas.ColWidth(1) = 2500
     FrmDatplac.Gridtablas.ColWidth(2) = 1
     FrmDatplac.Gridtablas.ColWidth(3) = 1
     FrmDatplac.Gridtablas.ColWidth(4) = 1
     FrmDatplac.Gridtablas.TextMatrix(0, 0) = "Item"
     FrmDatplac.Gridtablas.TextMatrix(0, 1) = "Producto"
     FrmDatplac.Gridtablas.TextMatrix(0, 2) = "Precio Desct."
     SQ_OPER = 2
     LEER_TAB_LLAVE
     FrmDatplac.Gridtablas.RowHeight(0) = 300
     FrmDatplac.Gridtablas.Rows = 1
     fila = 0
     Do Until tab_mayor.EOF
      fila = fila + 1
      FrmDatplac.Gridtablas.Rows = FrmDatplac.Gridtablas.Rows + 1
      FrmDatplac.Gridtablas.RowHeight(fila) = 280
      FrmDatplac.Gridtablas.TextMatrix(fila, 0) = Nulo_Valor0(tab_mayor!tab_numtab)
      FrmDatplac.Gridtablas.TextMatrix(fila, 1) = Trim(Nulo_Valors(tab_mayor!tab_nomlargo))
      FrmDatplac.Gridtablas.TextMatrix(fila, 2) = Nulo_Valors(tab_mayor!TAB_CODCIA)
      FrmDatplac.Gridtablas.TextMatrix(fila, 3) = Nulo_Valor0(tab_mayor!TAB_TIPREG)
      FrmDatplac.Gridtablas.TextMatrix(fila, 4) = Nulo_Valor0(tab_mayor!tab_numtab)
      tab_mayor.MoveNext
     Loop
     FrmDatplac.Gridtablas.Visible = True
     FrmDatplac.Gridtablas.SetFocus

End Sub

Private Sub Form_Load()
CenterMe FrmDatplac

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
    wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL) = Format(Val(wsTexto.Text), "0")
   Else
    wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL) = Format(Val(wsTexto.Text), "0.00")
   End If
   CmdAgregar.Enabled = True
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

Private Sub gridtablas_Click()
ElGrid_Click Gridtablas, textovar
End Sub

Private Sub gridtablas_DblClick()
If PUB_TIPREG = -55 Then Exit Sub
If Gridtablas.COL = 1 Then Exit Sub
ElGrid_DblClick Gridtablas, textovar
End Sub

Private Sub gridtablas_EnterCell()
ElGrid_EnterCell Gridtablas, textovar, 2, 3, 4
End Sub

Private Sub gridtablas_GotFocus()
ElGrid_GotFocus Gridtablas, textovar
End Sub

Private Sub gridtablas_KeyDown(KeyCode As Integer, Shift As Integer)
If PUB_TIPREG = -55 Then
  If Gridtablas.COL <> 3 Then Exit Sub
  If KeyCode = 45 Then Exit Sub
End If
If KeyCode = 13 And Gridtablas.COL = 1 Then
 Exit Sub
End If
If Gridtablas.Row = 0 And KeyCode <> 45 Then
 Exit Sub
End If
If KeyCode = 113 And Gridtablas.COL = 1 And PUB_TIPREG <> 2103 Then Exit Sub
If KeyCode = 45 Then
 cmdagregar_Click
 Exit Sub
End If
If KeyCode = 46 And PUB_TIPREG = -10 Then GoTo borra_transportista
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
    If PUB_TIPREG = 2301 Then
      PUB_CODCIA = LK_CODCIA
      PUB_NUMTAB = FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Row, 0)
    Else
      PUB_CODCIA = FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Row, 2)
      PUB_TIPREG = FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Row, 3)
      PUB_NUMTAB = FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Row, 4)
    End If
    
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
ElGrid_KeyDown Gridtablas, textovar, KeyCode
Exit Sub
borra_transportista:
pub_mensaje = " Eliminar    : " & Trim(Gridtablas.TextMatrix(Gridtablas.Row, 1)) & " ¿Desea Continuar... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
  Gridtablas.SetFocus
  Exit Sub
End If


Gridtablas.SetFocus

End Sub

Private Sub gridtablas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And PUB_TIPREG = 2301 Then
 gridtablas_KeyDown 113, 0
 Exit Sub
End If
If PUB_TIPREG = -55 And Gridtablas.COL = 3 Then
  ElGrid_KeyDown Gridtablas, textovar, 113
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
ElGrid_LeaveCell Gridtablas, textovar
End Sub

Private Sub gridtablas_Scroll()
If Gridtablas.RowIsVisible(Gridtablas.Row) = False Then
   textovar.Visible = False
   Exit Sub
Else
  textovar.Visible = False
End If

textovar.Left = Gridtablas.Left + Gridtablas.CellLeft ' gridtablas.CellLeft  'Frame4.Left + gridtablas.CellLeft
textovar.Width = Gridtablas.CellWidth
textovar.Top = Gridtablas.Top + Gridtablas.CellTop ' Frame4.Top + gridtablas.Top + gridtablas.CellTop
textovar.DataField = Gridtablas.COL
textovar.Tag = Gridtablas.Row
textovar.Text = Gridtablas.TextMatrix(Gridtablas.Row, Gridtablas.COL)

End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)

TEXTO_KeyPress Gridtablas, textovar, KeyAscii, 2, 5, 2, 6, 2

If KeyAscii = 13 Then
  If PUB_TIPREG = 2301 Then
     PSVE2_LLAVE(0) = LK_CODCIA
     PSVE2_LLAVE(1) = FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Row, 4)
     PSVE2_LLAVE(2) = FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Row, 7)
     clive2_llave.Requery
     If clive2_llave.EOF Then
       MsgBox "Reiniciar Sistema.", 48, Pub_Titulo
       Unload FrmDatplac
     End If
     clive2_llave.Edit
     clive2_llave!tab_nomlargo = FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Row, 2)
     clive2_llave!tab_nomcorto = FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Row, 3)
     clive2_llave.Update
    Exit Sub
  End If
  If PUB_TIPREG = -33 Then
     PSVE2_LLAVE(0) = LK_CODCIA
     PSVE2_LLAVE(1) = Val(FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Row, 0))
     clive2_llave.Requery
     If clive2_llave.EOF Then
       MsgBox "Descripción Existe. Intente Nuevamente ..", 48, Pub_Titulo
       Azul textovar, textovar
       Exit Sub
     End If
     clive2_llave.Edit
     clive2_llave!VE2_NOMBRE = FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Row, 1)
     clive2_llave.Update
   Exit Sub
  End If
  If PUB_TIPREG = 2103 Then
     SQ_OPER = 1
      PUB_CODCIA = FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Row, 2)
      PUB_TIPREG = FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Row, 3)
      PUB_NUMTAB = FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Row, 4)
      LEER_TAB_LLAVE
      If tab_llave.EOF Then
        Exit Sub
      End If
      tab_llave.Edit
      tab_llave!tab_nomlargo = FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Row, 1)
      tab_llave!tab_nomcorto = FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Row, 1)
      tab_llave!TAB_contable2 = 0
      tab_llave!tab_codART = 0
      tab_llave.Update
      Exit Sub
  End If
     'PSPLAC_LLAVE(0) = PUB_CODCIA
     'PSPLAC_LLAVE(1) = PUB_TIPREG
     'PSPLAC_LLAVE(2) = pu_codclie
     'PSPLAC_LLAVE(3) = textovar.Text 'FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Row, 4)
     'cliplac_llave.Requery
     'If Not cliplac_llave.EOF Then
     '  If cliplac_llave.RowCount > 1 Then
     '    MsgBox "Descripción Existe. Intente Nuevamente ..", 48, Pub_Titulo
     '    Azul textovar, textovar
     '    Exit Sub
     '  End If
     'End If
     'cliplac_llave.Edit
     'cliplac_llave!tab_nomlargo = FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Row, 1)
     'cliplac_llave!tab_nomcorto = FrmDatplac.Gridtablas.TextMatrix(Gridtablas.Row, 1)
     'cliplac_llave.Update
     
End If
Exit Sub

End Sub
Private Sub textovar_LostFocus()
If Gridtablas.Visible Then
 TEXTO_LosFocus Gridtablas, textovar
Else
 TEXTO_LosFocus grid_comi, textovar
End If
End Sub


' aqui enpi

Private Sub grid_comi_Click()
If grid_comi.Rows < 3 Then
 Exit Sub
End If
ElGrid_Click grid_comi, textovar
End Sub

Private Sub grid_comi_DblClick()
If grid_comi.Rows < 3 Then
 Exit Sub
End If
ElGrid_DblClick grid_comi, textovar
End Sub

Private Sub grid_comi_EnterCell()
If grid_comi.Rows < 3 Then
 Exit Sub
End If
ElGrid_EnterCell grid_comi, textovar
End Sub

Private Sub grid_comi_GotFocus()
ElGrid_GotFocus grid_comi, textovar
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
 If grid_comi.TextMatrix(FrmDatplac.grid_comi.Rows - 1, 1) = "0" And grid_comi.TextMatrix(FrmDatplac.grid_comi.Rows - 1, 2) = "0" And grid_comi.TextMatrix(FrmDatplac.grid_comi.Rows - 1, 3) = "0.00" Then
  MsgBox "No procede, se Agrego y esta en blanco. ", 48, Pub_Titulo
  Exit Sub
 End If
 FrmDatplac.grid_comi.Rows = FrmDatplac.grid_comi.Rows + 1
 FrmDatplac.grid_comi.RowHeight(FrmDatplac.grid_comi.Rows - 1) = 280
 FrmDatplac.grid_comi.TextMatrix(FrmDatplac.grid_comi.Rows - 1, 1) = "0"
 FrmDatplac.grid_comi.TextMatrix(FrmDatplac.grid_comi.Rows - 1, 2) = "0"
 FrmDatplac.grid_comi.TextMatrix(FrmDatplac.grid_comi.Rows - 1, 3) = "0.00"
 CmdAgregar.Enabled = True
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
    CmdAgregar.Enabled = True
    grid_comi.SetFocus
End If

ElGrid_KeyDown grid_comi, textovar, KeyCode

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
ElGrid_LeaveCell grid_comi, textovar
End Sub

Private Sub grid_comi_Scroll()
If grid_comi.RowIsVisible(grid_comi.Row) = False Then
   textovar.Visible = False
   Exit Sub
Else
  textovar.Visible = False
End If

textovar.Left = grid_comi.Left + grid_comi.CellLeft ' grid_comi.CellLeft  'Frame4.Left + grid_comi.CellLeft
textovar.Width = grid_comi.CellWidth
textovar.Top = grid_comi.Top + grid_comi.CellTop ' Frame4.Top + grid_comi.Top + grid_comi.CellTop
textovar.DataField = grid_comi.COL
textovar.Tag = grid_comi.Row
textovar.Text = grid_comi.TextMatrix(grid_comi.Row, grid_comi.COL)

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


Public Function BUSCA(WDATO As Currency) As Boolean
For fila = 1 To Gridtablas.Rows - 1
 If Val(FrmDatplac.Gridtablas.TextMatrix(fila, 6)) = WDATO Then
   BUSCA = True
   Exit Function
 End If
Next fila
BUSCA = False
End Function

