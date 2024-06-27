VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form centrocostos 
   BackColor       =   &H00FAEFDA&
   Caption         =   "Centros de Costos"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   Icon            =   "centrocostos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   6570
   Begin ComctlLib.ListView ListView2 
      Height          =   375
      Left            =   675
      TabIndex        =   13
      Top             =   1350
      Visible         =   0   'False
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   661
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FAEFDA&
      Height          =   780
      Left            =   15
      TabIndex        =   15
      Top             =   2910
      Width           =   6525
      Begin VB.CommandButton cmdcerrar 
         Caption         =   "Ce&rrar"
         Height          =   405
         Left            =   5040
         TabIndex        =   19
         Top             =   255
         Width           =   1215
      End
      Begin VB.CommandButton cmdcancelar 
         Caption         =   "&Cancelar"
         Height          =   405
         Left            =   3410
         TabIndex        =   18
         Top             =   255
         Width           =   1215
      End
      Begin VB.CommandButton cmdgrabar 
         Caption         =   "&Grabar"
         Height          =   405
         Left            =   150
         TabIndex        =   17
         Top             =   255
         Width           =   1215
      End
      Begin VB.CommandButton cmdeliminar 
         Caption         =   "&Eliminar"
         Height          =   405
         Left            =   1780
         TabIndex        =   16
         Top             =   255
         Width           =   1215
      End
   End
   Begin VB.Frame fracentros 
      BackColor       =   &H00FAEFDA&
      Height          =   2940
      Left            =   15
      TabIndex        =   6
      Top             =   -30
      Width           =   6540
      Begin VB.TextBox txtamarre3 
         Height          =   285
         Left            =   1665
         MaxLength       =   2
         TabIndex        =   5
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txtamarre2 
         Height          =   285
         Left            =   1665
         MaxLength       =   2
         TabIndex        =   4
         Top             =   2160
         Width           =   975
      End
      Begin VB.ComboBox CMBTIPO 
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
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   330
         Width           =   3615
      End
      Begin VB.TextBox txtamarre1 
         Height          =   285
         Left            =   1665
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtdescrip 
         Height          =   285
         Left            =   1665
         MaxLength       =   40
         TabIndex        =   2
         Top             =   1440
         Width           =   4455
      End
      Begin VB.TextBox txtcc 
         Height          =   285
         Left            =   1665
         MaxLength       =   4
         TabIndex        =   1
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblnumdgt 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4350
         TabIndex        =   21
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Num. de Digitos C.C. :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   2715
         TabIndex        =   20
         Top             =   1095
         Width           =   1560
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amarre a Clases. :"
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
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   1515
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amarre a Clases. :"
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
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   1515
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cent. C. :"
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
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amarre a Clases. :"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion C.C. :"
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
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo C.C. :"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1005
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
      Left            =   -15
      TabIndex        =   14
      Top             =   3720
      Width           =   6615
   End
End
Attribute VB_Name = "centrocostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim loc_key As Integer
Dim PSCC_ACTIVO As rdoQuery
Dim cc_activo As rdoResultset


Private Sub CMBTIPO_Click()
    PUB_TIPREG = Val(Left(CMBTIPO.Text, 2))
    If PUB_TIPREG = 1 Then
        lblnumdgt = PUB_Dgt_CC1
    ElseIf PUB_TIPREG = 2 Then
        lblnumdgt = PUB_Dgt_CC2
    ElseIf PUB_TIPREG = 3 Then
        lblnumdgt = PUB_Dgt_CC3
    End If
End Sub

Private Sub CMBTIPO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Azul txtcc, txtcc
End Sub

Private Sub cmdcancelar_Click()
    Limpia_centros
End Sub

Private Sub cmdcerrar_Click()
    Unload centrocostos
End Sub

Private Sub cmdeliminar_Click()
If Val(PUB_TIPREG) = 0 Then GoTo NOPRO
If Trim(txtcc) = "" Then GoTo NOPRO

pub_mensaje = "Confirmar la Eliminación del Registro... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
Else

End If
cc_activo.Requery
If cc_activo.EOF Then
  GoTo NOPRO
Else
 cc_activo.Delete
End If
Limpia_centros

Exit Sub
NOPRO:
 MsgBox "NO procede,verificar ", 48, Pub_Titulo


End Sub

Private Sub cmdgrabar_Click()
If Val(PUB_TIPREG) = 0 Then GoTo NOPRO

If Trim(txtcc) = "" Then GoTo NOPRO


If cc_activo.EOF Then
 cc_activo.AddNew
Else
 cc_activo.Edit
End If
cc_activo!cc_codcia = LK_CODCIA
cc_activo!cc_TIPO = PUB_TIPREG
cc_activo!cc_codigo = Trim(txtcc.Text)


cc_activo!cc_descripcion = txtdescrip.Text
'cc_activo!cc_amarre1 = Trim(txtamarre1.Text)
'cc_activo!cc_amarre2 = Trim(txtamarre2.Text)
'cc_activo!cc_amarre3 = Trim(txtamarre3.Text)
cc_activo.Update
Limpia_centros

Exit Sub
NOPRO:
 MsgBox "NO procede,verificar ", 48, Pub_Titulo


End Sub

Private Sub Form_Activate()
If CMBTIPO.ListCount > 0 Then CMBTIPO.ListIndex = 0
End Sub

Private Sub Form_Load()
CenterMe centrocostos
pub_cadena = "SELECT * FROM CENTROC WHERE CC_CODCIA = ? AND CC_TIPO = ? AND CC_CODIGO = ? "
Set PSCC_ACTIVO = CN.CreateQuery("", pub_cadena)
PSCC_ACTIVO(0) = 0
PSCC_ACTIVO(1) = 0
PSCC_ACTIVO(2) = 0
Set cc_activo = PSCC_ACTIVO.OpenResultset(rdOpenKeyset, rdConcurValues)

CMBTIPO.AddItem "01.- ADMINISTRACION - 500"
CMBTIPO.AddItem "02.- VENTAS - 600"
CMBTIPO.AddItem "03.- FINANZAS - 700"

End Sub

Private Sub txtamarre1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtamarre2.SetFocus
End Sub

Private Sub txtamarre2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtamarre3.SetFocus
End Sub

Private Sub txtamarre3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If cmdgrabar.Enabled Then cmdgrabar.SetFocus
End If
End Sub

Private Sub txtcc_GotFocus()
'Azul txtcc, txtcc

End Sub
Private Sub txtcc_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not ListView2.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And txtcc.Text = "" Then
  loc_key = 1
  Set ListView2.SelectedItem = ListView2.ListItems(loc_key)
  ListView2.ListItems.Item(loc_key).Selected = True
  ListView2.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView2.ListItems.Count Then loc_key = ListView2.ListItems.Count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView2.ListItems.Count Then loc_key = ListView2.ListItems.Count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
'  KeyCode = 0
  ListView2.ListItems.Item(loc_key).Selected = True
  ListView2.ListItems.Item(loc_key).EnsureVisible
  txtcc.Text = Trim(ListView2.ListItems.Item(loc_key).Text) & " "
  DoEvents
  txtcc.SelStart = Len(txtcc.Text)
  DoEvents
fin:

End Sub
Private Sub txtcc_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As ListItem    ' Variable FoundItem.
If KeyAscii = 27 Then
 ListView2.Visible = False
 txtcc.Text = ""
 Exit Sub
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
On Error GoTo ERROR_CODIGO
pu_codclie = Val(txtcc.Text)
On Error GoTo 0
If Len(txtcc.Text) = 0 Then
   Exit Sub
End If

If pu_codclie <> 0 And IsNumeric(txtcc.Text) = True Then
    PSCC_ACTIVO(0) = LK_CODCIA
    PSCC_ACTIVO(1) = PUB_TIPREG
    PSCC_ACTIVO(2) = txtcc.Text
    cc_activo.Requery
    If cc_activo.EOF Then
       pub_mensaje = "No Existe Centro de Costo !!! ...  ¿Desea Adicionar... ?"
       Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
       If Pub_Respuesta = vbNo Then
          txtcc.Text = ""
          Exit Sub
       Else
        txtdescrip.SetFocus
       End If
    Else
       llena_centros
       Azul txtdescrip, txtdescrip
    End If

Else
   If loc_key > ListView2.ListItems.Count Or loc_key = 0 Then
     Exit Sub
   End If
   valor = UCase(ListView2.ListItems.Item(loc_key).Text)
   If Trim(UCase(txtcc.Text)) = Left(valor, Len(Trim(txtcc.Text))) Then
   Else
      Exit Sub
   End If
   'lblCliente.Caption = Trim(ListView2.ListItems.Item(loc_key).Text)
   txtcc.Text = Trim(ListView2.ListItems.Item(loc_key).SubItems(1))
   txtcc_KeyPress 13
   
   'If pantalla.Visible And pantalla.Enabled Then
   '  pantalla.SetFocus
   'End If
End If

dale:
ListView2.Visible = False
fin:
Exit Sub
ERROR_CODIGO:
MsgBox "Codigo NO Valido .... ", 48, Pub_Titulo
Azul txtcc, txtcc

End Sub

Private Sub txtcc_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
If Len(txtcc.Text) = 0 Or IsNumeric(txtcc.Text) = True Then
   ListView2.Visible = False
   Exit Sub
End If
If ListView2.Visible = False And KeyCode <> 13 Then
    var = Asc(txtcc.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    ElseIf var = 58 Then
       var = "A"
    Else
       var = Chr(var)
    End If
    numarchi = 6
    PUB_CP = "C"
    
    'archi = "SELECT CLI_CODCLIE, CLI_CODCIA, CLI_CP, CLI_NOMBRE,CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM  FROM CLIENTES WHERE  CLI_CP = '" & PUB_CP & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & txtcc.Text & "' AND  '" & VAR & "' ORDER BY CLI_NOMBRE"
    archi = "SELECT CC_CODIGO , CC_CODCIA, CC_DESCRIPCION FROM CENTROC WHERE CC_CODCIA = '" & LK_CODCIA & "'  AND CC_TIPO = " & PUB_TIPREG & " AND  CC_DESCRIPCION BETWEEN '" & txtcc.Text & "' AND  '" & var & "' ORDER BY CC_DESCRIPCION"
'    If Trim(txtcc.text) <> "" And ListView1.ListItems.count = 0 Then
'    Else
     PROC_LISVIEW ListView2
    ListView2.Top = 1350
    ListView2.Left = 1250
     loc_key = 0
     If ListView2.Visible Then
      loc_key = 1
     End If
 '   End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If ListView2.Visible Then
  Set itmFound = ListView2.FindItem(LTrim(txtcc.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView2.ListItems.Count Then
      ListView2.ListItems.Item(ListView2.ListItems.Count).EnsureVisible
   Else
     ListView2.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If


End Sub


Public Sub llena_centros()
If cc_activo.EOF Then Exit Sub
txtcc.Text = Trim(cc_activo!cc_codigo)
txtdescrip.Text = cc_activo!cc_descripcion
txtdescrip.Text = cc_activo!cc_descripcion
'bloqueado por mic campos no existen!!!!!!!!!!!!!!
'txtamarre1.Text = cc_activo!cc_amarre1
'txtamarre2.Text = cc_activo!cc_amarre2
'txtamarre3.Text = cc_activo!cc_amarre3

End Sub
Public Sub Limpia_centros()
txtcc.Text = ""
txtdescrip.Text = ""
txtdescrip.Text = ""
txtamarre1.Text = ""
txtamarre2.Text = ""
txtamarre3.Text = ""
If txtcc.Enabled Then txtcc.SetFocus
End Sub

Private Sub txtcc_LostFocus()
    If Len(txtcc.Text) <> Val(lblnumdgt) And Trim(txtcc.Text) <> "" Then
        MsgBox "Digite el Numero de Digitos Correcto"
        txtcc.SetFocus
        Azul txtcc, txtcc
    End If
    txtcc_KeyPress 13
End Sub

Private Sub txtdescrip_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtamarre1.SetFocus
End Sub
