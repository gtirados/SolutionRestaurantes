VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "Comctl32.ocx"
Begin VB.Form frmStockG 
   Caption         =   "Stock por Regla."
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdeliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   2400
      TabIndex        =   17
      Top             =   3720
      Width           =   1215
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   4200
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Ce&rrar"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdmodificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos :"
      Height          =   3375
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtstockr 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3720
         TabIndex        =   4
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txt_key 
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtnumfac 
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtturno 
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   1320
         Width           =   615
      End
      Begin MSMask.MaskEdBox txtfecha 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label lblturno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label nomarti 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   14
         Top             =   1800
         Width           =   4095
      End
      Begin VB.Label lblparte 
         Caption         =   "Stock según regla :"
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
         Index           =   4
         Left            =   1800
         TabIndex        =   13
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label lblparte 
         Caption         =   "Tanque/Producto :"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   12
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblparte 
         Caption         =   "Fecha :"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblparte 
         Caption         =   "Nº Documento:"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblparte 
         Caption         =   "Turno :"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmStockG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim loc_key As Integer
Dim stock_llave As rdoResultset
Dim PSST_LLAVE As rdoQuery
Dim FLAG_NUEVO   As String * 1
Dim VAR_ACTIVAR As Integer
Dim WCOD_ORIGINAL As Currency

Private Sub cmdcancelar_Click()
LIMPIA
OBBLO
cmdmodificar.Caption = "&Modificar"
FLAG_NUEVO = ""
txtnumfac.SetFocus
End Sub

Private Sub cmdCerrar_Click()
Unload frmStockG
End Sub

Private Sub cmdeliminar_Click()
If Val(txtnumfac.Text) <= 0 Then
 txtnumfac.SetFocus
 Exit Sub
End If
pub_mensaje = "Desea Eliminar este Registro... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If

SQ_OPER = 1
PUB_TIPREG = 2110
PUB_NUMTAB = Val(txtnumfac.Text)
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_llave.EOF Then Exit Sub
tab_llave.Delete
LIMPIA
OBBLO
cmdcancelar_Click
End Sub

Private Sub cmdmodificar_Click()
If Val(txtnumfac.Text) <= 0 Then
 txtnumfac.SetFocus
 Exit Sub
End If
If Left(cmdmodificar.Caption, 2) = "&M" Then
'    cmdeliminar.Enabled = False
    cmdmodificar.Caption = "&Grabar"
    OBDES
    txtfecha.SetFocus
Else
   SQ_OPER = 1
   PUB_TIPREG = 2102
   PUB_NUMTAB = Val(txtturno.Text)
   PUB_CODCIA = LK_CODCIA
   LEER_TAB_LLAVE
   If tab_llave.EOF Then
     Azul txtturno, txtturno
     MsgBox "Turno NO Existe...", 48, Pub_Titulo
     txtturno.SetFocus
     Exit Sub
   End If
   SQ_OPER = 3
    pu_alterno = Trim(txt_key.Text)
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_llave_alt.EOF Then
      MsgBox "Codigo del producto No Exiete ", 48, Pub_Titulo
      txt_key.SetFocus
      Exit Sub
    End If

  PSST_LLAVE(0) = LK_CODCIA
  PSST_LLAVE(1) = 2110
  PSST_LLAVE(2) = Format(txtfecha.Text, "dd/mm/yyyy")
  PSST_LLAVE(3) = Val(txtturno.Text)
  If FLAG_NUEVO = "A" Then
    PSST_LLAVE(4) = -1
   Else
   PSST_LLAVE(4) = Val(txtnumfac.Text)
  End If
  stock_llave.Requery
  If Not stock_llave.EOF Then
     MsgBox "Kardex ya Existe Pertenece a: " & Chr(13) & "Turno: " & stock_llave!tab_codclie & Chr(13) & "Fecha: " & stock_llave!tab_nomcorto
     Exit Sub
 End If
 
SQ_OPER = 1
PUB_TIPREG = 2110
PUB_NUMTAB = Val(txtnumfac.Text)
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
 
 cmdmodificar.Caption = "&Modificar"
' cmdeliminar.Enabled = True
 GRABA_STOCK
 OBBLO
 LIMPIA
' cmdmostrar.Enabled = True
  FLAG_NUEVO = ""
 txtnumfac.SetFocus
End If




End Sub

Private Sub Form_Load()
CenterMe frmStockG
pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_TIPREG = ?  AND TAB_NOMCORTO = ? AND TAB_CODCLIE = ? AND TAB_NUMTAB <> ? AND TAB_CODART = ?  ORDER BY TAB_NUMTAB "
Set PSST_LLAVE = CN.CreateQuery("", pub_cadena)
PSST_LLAVE(0) = 0
PSST_LLAVE(1) = 0
PSST_LLAVE(2) = 0
PSST_LLAVE(3) = 0
PSST_LLAVE(4) = 0
PSST_LLAVE(5) = 0
Set stock_llave = PSST_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
txtfecha.Visible = True
txtfecha.Mask = "##/##/####"
txtfecha.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
FLAG_NUEVO = ""
VAR_ACTIVAR = 0
OBBLO
End Sub

Public Sub OBBLO()
txtturno.Locked = True
txtturno.BackColor = QBColor(7)
txt_key.Locked = True
txt_key.BackColor = QBColor(7)
txtfecha.Enabled = False
txtstockr.Locked = True
txtstockr.BackColor = QBColor(7)
End Sub
Public Sub OBDES()
txtturno.Locked = False
txtturno.BackColor = QBColor(15)
txt_key.Locked = False
txt_key.BackColor = QBColor(15)
txtfecha.Enabled = True
txtstockr.Locked = False
txtstockr.BackColor = QBColor(15)
End Sub
Public Sub LIMPIA()
nomarti.Caption = ""
txtnumfac.Text = ""
txtturno.Text = ""
txt_key.Text = ""
txtfecha.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
txtstockr.Text = ""
txtstockr.Text = Format(Val(txtstockr.Text), "0.00")
 txtnumfac.Enabled = True

End Sub

Private Sub txtfecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If Not IsDate(txtfecha) Then
    MsgBox "Fecha no procede.", 48, Pub_Titulo
    Azul2 txtfecha, txtfecha
   Exit Sub
 End If
 Azul txtturno, txtturno
End If

End Sub

Private Sub txtnumfac_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii <> 13 Then Exit Sub
   If Val(txtnumfac.Text) <= 0 Then Exit Sub
   SQ_OPER = 1
   PUB_TIPREG = 2110
   PUB_NUMTAB = Val(txtnumfac.Text)
   PUB_CODCIA = LK_CODCIA
   LEER_TAB_LLAVE
   If tab_llave.EOF Then
     pub_mensaje = "No Existe Kardex. Desea Agregar uno Nuevo?"
     Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
     If Pub_Respuesta = vbNo Then
        Azul txtnumfac, txtnumfac
        Exit Sub
     End If
     FLAG_NUEVO = "A"
     txtnumfac.Enabled = False
     OBDES
     cmdmodificar.Caption = "&Grabar Kardex"
     cmdmodificar.Enabled = True
     Azul2 txtfecha, txtfecha
     Exit Sub
   End If
   LLENA
   cmdcancelar.SetFocus
End Sub

Public Sub LLENA()
If tab_llave.EOF Then Exit Sub
txtnumfac.Text = tab_llave!tab_numtab
txtturno.Text = tab_llave!tab_codclie
txtstockr.Text = Trim(tab_llave!tab_nomlargo)
txtfecha.Text = Format(tab_llave!tab_nomcorto, "dd/mm/yyyy")
SQ_OPER = 1
pu_codcia = LK_CODCIA
PUB_KEY = tab_llave!tab_codART
LEER_ART_LLAVE
If Not art_LLAVE.EOF Then
 nomarti.Caption = art_LLAVE!art_nombre
 txt_key.Text = art_LLAVE!art_alterno
End If

 
End Sub

Private Sub txtstockr_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL txtstockr, KeyAscii
If KeyAscii <> 13 Then Exit Sub
  cmdmodificar.SetFocus
End Sub

Private Sub txtturno_Change()
lblturno.Caption = ""
End Sub

Private Sub txtturno_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 27 Then
  txtturno.Text = ""
  Exit Sub
End If
If KeyAscii <> 13 Then Exit Sub
   SQ_OPER = 1
   PUB_TIPREG = 2102
   PUB_NUMTAB = Val(txtturno.Text)
   PUB_CODCIA = LK_CODCIA
   LEER_TAB_LLAVE
   If tab_llave.EOF Then
     Azul txtturno, txtturno
     MsgBox "Turno NO Existe...", 48, Pub_Titulo
     txtturno.SetFocus
     Exit Sub
   End If
   txtturno.Text = PUB_NUMTAB
   lblturno.Caption = tab_llave!tab_nomlargo
 
 Azul txt_key, txt_key


'PSST_LLAVE(0) = LK_CODCIA
'PSST_LLAVE(1) = 0
'PSST_LLAVE(2) = 0
'stock_llave.Requery


End Sub

Public Sub GRABA_STOCK()
If tab_llave.EOF Then
  tab_llave.AddNew
Else
  tab_llave.Edit
End If
    tab_llave!TAB_CODCIA = LK_CODCIA
    tab_llave!tab_numtab = Val(txtnumfac.Text)
    tab_llave!TAB_TIPREG = 2110
    tab_llave!tab_codclie = Val(txtturno.Text)
    tab_llave!tab_nomlargo = Format(txtstockr.Text, "0.00")
    tab_llave!tab_nomcorto = Format(txtfecha.Text, "dd/mm/yyyy")
    SQ_OPER = 3
    pu_alterno = Trim(txt_key.Text)
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_llave_alt.EOF Then
      MsgBox "Codigo No Exiete "
      Unload frmStockG
      Exit Sub
    End If
    tab_llave!tab_codART = art_llave_alt!ART_KEY
tab_llave.Update

End Sub

Private Sub txt_key_Change()
If txt_key = "" Then
 nomarti.Caption = ""
End If
End Sub

Private Sub txt_key_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And txt_key.Text = "" Then
  loc_key = 1
  Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  GoTo FIN
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo FIN
POSICION:
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text)
  DoEvents
  txt_key.SelStart = Len(txt_key.Text)
  DoEvents
FIN:

End Sub
Private Sub txt_key_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i, car
Dim itmFound As ListItem
car = Chr(KeyAscii)
KeyAscii = Asc(UCase(car))
If KeyAscii = 27 Then
 ListView1.Visible = False
 txt_key.Text = ""
End If
If KeyAscii <> 13 Then
   GoTo FIN
End If
VAR_ACTIVAR = 0
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  PUB_KEY = 0
Else
 PUB_KEY = Val(txt_key.Text)
 If Len(txt_key.Text) = 0 Then
    Exit Sub
 End If
 If IsNumeric(txt_key.Text) = False Then
   PUB_KEY = 0
 End If
End If

If PUB_KEY <> 0 Then
    SQ_OPER = 1
    PUB_KEY = txt_key.Text
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then
       MsgBox "Codigo NO Existe.", 48, Pub_Titulo
       Azul txt_key, txt_key
       GoTo FIN
    End If
    WCOD_ORIGINAL = art_LLAVE!ART_KEY
    nomarti.Caption = Trim(art_LLAVE!art_nombre)
    GoSub VERI
    ListView1.Visible = False
    Azul txtstockr, txtstockr
    Exit Sub
Else
  If ListView1.Visible = False And VAR_ACTIVAR <> 99 And txt_key.Text <> "" And LK_FLAG_ORIGINAL <> "A" And LK_FLAG_ALTERNO = "A" Then
IR_ALTERNO:
     SQ_OPER = 3
     pu_alterno = txt_key.Text
     pu_codcia = LK_CODCIA
     LEER_ART_LLAVE
     If art_llave_alt.EOF Then
       MsgBox "Codigo No Existe ...", 48, Pub_Titulo
       Azul txt_key, txt_key
       Exit Sub
     End If
     WCOD_ORIGINAL = art_llave_alt!ART_KEY
     nomarti.Caption = Trim(art_llave_alt!art_nombre)
     GoSub VERI
     ListView1.Visible = False
     Azul txtstockr, txtstockr
     Exit Sub
  Else
    If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
    End If
    valor = UCase(ListView1.ListItems.Item(loc_key).Text)
    If Trim(UCase(txt_key.Text)) = Left(valor, Len(Trim(txt_key.Text))) And Len(Trim(txt_key.Text)) <> 0 Then
      If VAR_ACTIVAR = 0 And LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
        txt_key.Text = Trim(ListView1.ListItems.Item(loc_key))
        GoTo IR_ALTERNO
      End If
      If VAR_ACTIVAR <> 99 Then
       txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
      Else
       txt_key.Text = Trim(ListView1.ListItems.Item(loc_key))
      End If
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       PUB_KEY = Val(ListView1.ListItems.Item(loc_key).SubItems(1))
      Else
       PUB_KEY = txt_key.Text
      End If
      LEER_ART_LLAVE
      VAR_ACTIVAR = 0
      If art_LLAVE.EOF Then
        MsgBox "Codigo No Existe ...", 48, Pub_Titulo
        Azul txt_key, txt_key
        Exit Sub
      End If
      WCOD_ORIGINAL = art_LLAVE!ART_KEY
      nomarti.Caption = Trim(art_LLAVE!art_nombre)
      GoSub VERI
      Azul txtstockr, txtstockr
      ListView1.Visible = False
      Exit Sub
    Else
      Exit Sub
    End If
    
  End If
End If
dale:
ListView1.Visible = False
FIN:
Exit Sub
ERROR_CODIGO:
MsgBox "Codigo NO Valido .... ", 48, Pub_Titulo
Azul txt_key, txt_key
Exit Sub

VERI:
  PSST_LLAVE(0) = LK_CODCIA
  PSST_LLAVE(1) = 2110
  PSST_LLAVE(2) = Format(txtfecha.Text, "dd/mm/yyyy")
  PSST_LLAVE(3) = Val(txtturno.Text)
  If FLAG_NUEVO = "A" Then
    PSST_LLAVE(4) = -1
  Else
    PSST_LLAVE(4) = Val(txtnumfac.Text)
  End If
  PSST_LLAVE(5) = WCOD_ORIGINAL
  stock_llave.Requery
  If Not stock_llave.EOF Then
     MsgBox "Kardex ya Existe Pertenece a: " & Chr(13) & "Nº.Kardex: " & stock_llave!tab_numtab & Chr(13) & "Turno: " & stock_llave!tab_codclie & Chr(13) & "Fecha: " & stock_llave!tab_nomcorto & Chr(13) & "Tanque/Producto: " & Trim(txt_key.Text) & " " & Trim(nomarti.Caption) & Chr(13) & "Stock x Regla.: " & Format(Val(stock_llave!tab_nomlargo), "#,##0.00"), 48, Pub_Titulo
     txt_key.Text = ""
     txt_key.SetFocus
     Exit Sub
     
 End If
 Return

End Sub

Private Sub txt_key_KeyUp(KeyCode As Integer, Shift As Integer)
Dim VAR
Dim ws_codcia As String * 2
Dim wnormal As String
wnormal = "1"
If KeyCode = 13 Then Exit Sub
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  If Len(txt_key.Text) = 0 Or txt_key.Text = "" Then
    ListView1.Visible = False
    Exit Sub
  End If
  If txt_key.Text = "*" And KeyCode = 106 Then
   VAR_ACTIVAR = 99
   Exit Sub
  ElseIf txt_key.Text = "" Then
   VAR_ACTIVAR = 0
   Exit Sub
  End If
  If VAR_ACTIVAR <> 99 Then
    Exit Sub
  End If
  If Left(txt_key.Text, 1) = "*" Then
   txt_key.Text = Mid(txt_key.Text, 2, Len(txt_key.Text))
   txt_key.SelStart = Len(txt_key.Text)
  End If
Else
 If Len(txt_key.Text) = 0 Or IsNumeric(txt_key.Text) = True Then
   ListView1.Visible = False
   Exit Sub
 End If
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(txt_key.Text) = 1 Then
    VAR = Asc(txt_key.Text)
    VAR = VAR + 1
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    ws_codcia = LK_CODCIA
    If LK_EMP_PTO = "A" Then
      ws_codcia = "00"
    End If
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      numarchi = 3
      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_KEY <> 0 AND ART_CALIDAD = " & wnormal & " AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & txt_key.Text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
    Else
      numarchi = 0
      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE ART_CALIDAD = " & wnormal & " AND  ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & txt_key.Text & "' AND  '" & VAR & "' ORDER BY ART_NOMBRE"
    End If
     PROC_LISVIEW ListView1
     loc_key = 0
     If ListView1.Visible Then
      loc_key = 1
     End If
     Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(txt_key.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView1.ListItems.count Then
      ListView1.ListItems.Item(ListView1.ListItems.count).EnsureVisible
   Else
     ListView1.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If

End Sub

