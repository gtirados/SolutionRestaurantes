VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmBuscacta 
   Caption         =   "Busqueda de Cuentas Contables"
   ClientHeight    =   4560
   ClientLeft      =   4740
   ClientTop       =   1680
   ClientWidth     =   4425
   Icon            =   "FRMBUSCACTA.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   4425
   Begin VB.TextBox txtdatos 
      Height          =   285
      Index           =   7
      Left            =   5160
      TabIndex        =   16
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtdatos 
      Height          =   285
      Index           =   6
      Left            =   5160
      TabIndex        =   15
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtdatos 
      Height          =   285
      Index           =   5
      Left            =   5160
      TabIndex        =   14
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtdatos 
      Height          =   285
      Index           =   4
      Left            =   5160
      TabIndex        =   13
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtdatos 
      Height          =   285
      Index           =   3
      Left            =   5160
      TabIndex        =   12
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtdatos 
      Height          =   285
      Index           =   2
      Left            =   5160
      TabIndex        =   11
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtdatos 
      Height          =   285
      Index           =   1
      Left            =   5160
      TabIndex        =   10
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtdatos 
      Height          =   285
      Index           =   0
      Left            =   5160
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdcolocar 
      Caption         =   "&Aceptar"
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
      Left            =   360
      TabIndex        =   5
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Plan Contable :"
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   4335
      Begin ComctlLib.TreeView TreeView1 
         Height          =   2535
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   4471
         _Version        =   327682
         HideSelection   =   0   'False
         Style           =   6
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox tnombre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   7
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox tcuenta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   6
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label etiqueta 
         Caption         =   "Cuenta y Descripción para Colocar :"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2880
         Width           =   2895
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000040C0&
      BorderWidth     =   10
      X1              =   0
      X2              =   9480
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lbltabla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Cuenta Contable para :"
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
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmBuscacta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim com_cont As rdoResultset
Dim PSCOM_CONT As rdoQuery
Dim PSCOM_MAYOR2 As rdoQuery
Dim com_mayor2 As rdoResultset
Dim NIVEL_ACT As Integer
Dim CARAC
Dim fila As Integer
Dim wCOM_NIVEL(6) As Integer
Dim NIVEL_MAX As Integer
Dim ww_nivel As Integer
Dim qq_cuenta As String * 12
Dim qq_indice As String * 12
Dim Posi_Ultimo As Integer
Dim Modo_Acceso As String

Public Function BUSCA_AUTO(valor As String) As Boolean
Dim ba As rdoResultset
Dim cade As String
  cade = "SELECT * FROM COMAEST WHERE COM_CUENTA = '" & Trim(valor) & "' AND COM_CODCIA = '" & LK_CODCIA & "' ORDER BY COM_CUENTA"
  Set ba = CN.OpenResultset(cade, rdOpenKeyset, rdConcurValues)
   ba.Requery
   BUSCA_AUTO = False
   If ba.EOF Then
     Exit Function
   End If
   If ba!com_FLAG_AFECTACION = "1" Then
     BUSCA_AUTO = True
   End If

End Function

Public Sub VERI_NIVEL(valor As String)
Dim cade As String
Dim wnivel As Integer
Dim i, windex As Integer
Dim wBusca As String

LOC_CTA_SUP = ""
NIVEL_ACT = 0
cade = Trim(valor)
wnivel = Len(cade)
windex = 0
For i = 1 To 6
  If wCOM_NIVEL(i) = wnivel Then
     windex = i
     Exit For
  End If
Next i
If windex = 0 Then
    MsgBox "Nro. de digitos invalidos, Verificar ..", 48, Pub_Titulo
    NIVEL_ACT = 0
   Exit Sub
End If
If windex <> 1 Then
'SQ_OPER = 1
'PUB_CUENTA = grid_cont.TextMatrix(fila - 1, 2)
'LEER_COM_LLAVE
'If com_llave.EOF Then
'   MsgBox "Cuenta Superior NO Existe, Verificar ...", 48, Pub_Titulo
'   Exit Sub
'End If
End If
NIVEL_ACT = windex
End Sub



Private Sub cmdcerrar_Click()
tcuenta.text = ""
tnombre.text = ""
frmBuscacta.Hide
End Sub

Private Sub cmdcolocar_Click()
frmBuscacta.Hide
End Sub

Private Sub Form_Load()
Dim barrita As ProgressBar
Dim i
Dim ws_tipo_cta As Integer
Dim WS_SIGNO_D, WS_SIGNO_H As Integer
Dim nodX As Node
On Error GoTo sigue:
Modo_REICIO = ""
Modo_Acceso = ""
Posi_Ultimo = -1
pub_cadena = "SELECT * FROM COPARAM WHERE COP_CODCIA = ?"
Set PSCOP_LLAVE = CN.CreateQuery("", pub_cadena)
PSCOP_LLAVE(0) = 0
Set cop_llave = PSCOP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

PSCOP_LLAVE.rdoParameters(0) = LK_CODCIA
cop_llave.Requery
If cop_llave.EOF Then
 Screen.MousePointer = 0
 MsgBox "Hay que definir parametros de contabilidad.", 48, Pub_Titulo
 Unload FORM_CONTA
 Exit Sub
End If


If Not cop_llave.EOF Then
For i = 1 To 6
  If cop_llave.rdoColumns(i) <> 0 Then
     wCOM_NIVEL(i) = cop_llave.rdoColumns(i)
     NIVEL_MAX = i
  End If
Next i
Else
 MsgBox "Definir parametros para el plan contable.", 48, Pub_Titulo
 Exit Sub
End If
'archi viene de cada tabla
Set PSCOM_CONT = CN.CreateQuery("", archi)
PSCOM_CONT(0) = 0
Set com_cont = PSCOM_CONT.OpenResultset(rdOpenKeyset, rdConcurValues)
If LK_EMP_PTO = "A" Then
 PSCOM_CONT(0) = "00"
Else
 PSCOM_CONT(0) = LK_CODCIA
End If
com_cont.Requery
fila = 0
Set nodX = TreeView1.Nodes.Add(, tvwChild, "TITULO", "Cuentas Principales")
TreeView1.Nodes.Item(1).Tag = "RAIZ"
TreeView1.Nodes.Item(1).Expanded = True
TreeView1.Indentation = 400
If Trim(LK_TABLA) = "CLIENTE" Or Trim(LK_TABLA) = "PROVEEDOR" Then
  Set barrita = frmCLI.PB
ElseIf Trim(LK_TABLA) = "PROVEEDOR2" Or Trim(LK_TABLA) = "CLIENTES2" Then
  Set barrita = frmCLI.PB2
ElseIf Trim(LK_TABLA) = "ARTI1" Then
  Set barrita = frmARTI.PB1
ElseIf Trim(LK_TABLA) = "ARTI2" Then
  Set barrita = frmARTI.PB2
ElseIf Trim(LK_TABLA) = "ARTI3" Then
  Set barrita = frmARTI.PB3
ElseIf Trim(LK_TABLA) = "ARTI4" Then
  Set barrita = frmARTI.PB4
ElseIf Trim(LK_TABLA) = "BANCO" Then
  Set barrita = frmBancos.PB
ElseIf LK_TABLA = "BUSCAR" Then
  Set barrita = FORM_CONTA.PB
ElseIf LK_TABLA = "SUB_TRA" Then
   Set barrita = FrmSub.bp
End If
barrita.Visible = True
barrita.Value = 0
barrita.Min = 0
barrita.Max = com_cont.RowCount
Do Until com_cont.EOF
   barrita.Value = barrita.Value + 1
   fila = fila + 1
   If com_cont!COM_NIVEL = 1 Then
      Set nodX = TreeView1.Nodes.Add("TITULO", tvwChild, "A" + Left(com_cont!com_cuenta, wCOM_NIVEL(1)), Trim(com_cont!com_cuenta) + " " + Trim(com_cont!com_DESCRIPCION))
   ElseIf com_cont!COM_NIVEL = 2 Then
      Set nodX = TreeView1.Nodes.Add("A" + Left(com_cont!com_cuenta, wCOM_NIVEL(1)), tvwChild, "B" + Left(com_cont!com_cuenta, wCOM_NIVEL(2)), Trim(com_cont!com_cuenta) + "  " + Trim(com_cont!com_DESCRIPCION))
   ElseIf com_cont!COM_NIVEL = 3 Then
      Set nodX = TreeView1.Nodes.Add("B" + Left(com_cont!com_cuenta, wCOM_NIVEL(2)), tvwChild, "C" + Left(com_cont!com_cuenta, wCOM_NIVEL(3)), Trim(com_cont!com_cuenta) + "  " + Trim(com_cont!com_DESCRIPCION))
   ElseIf com_cont!COM_NIVEL = 4 Then
      Set nodX = TreeView1.Nodes.Add("C" + Left(com_cont!com_cuenta, wCOM_NIVEL(3)), tvwChild, "D" + Left(com_cont!com_cuenta, wCOM_NIVEL(4)), Trim(com_cont!com_cuenta) + "  " + Trim(com_cont!com_DESCRIPCION))
   ElseIf com_cont!COM_NIVEL = 5 Then
      Set nodX = TreeView1.Nodes.Add("D" + Left(com_cont!com_cuenta, wCOM_NIVEL(4)), tvwChild, "E" + Left(com_cont!com_cuenta, wCOM_NIVEL(5)), Trim(com_cont!com_cuenta) + "  " + Trim(com_cont!com_DESCRIPCION))
   ElseIf com_cont!COM_NIVEL = 6 Then
      Set nodX = TreeView1.Nodes.Add("E" + Left(com_cont!com_cuenta, wCOM_NIVEL(5)), tvwChild, "F" + Left(com_cont!com_cuenta, wCOM_NIVEL(6)), Trim(com_cont!com_cuenta) + "  " + Trim(com_cont!com_DESCRIPCION))
   End If
   TreeView1.Nodes.Item(fila + 1).Tag = Str(com_cont!COM_NIVEL)
   TreeView1.Nodes.Item(fila + 1).Sorted = True
   If com_cont!COM_NIVEL = NIVEL_MAX - 2 Then
      TreeView1.Nodes.Item(fila + 1).Expanded = True
   End If
   If com_cont!COM_NIVEL = NIVEL_MAX - 1 Then
      txtdatos(0).text = com_cont!COM_NIVEL + 1
      txtdatos(1).text = com_cont!com_cuenta ' SUPERIOR
      txtdatos(2).text = "1" 'com_cont!COM_FLAG_AFECTACION
      txtdatos(3).text = com_cont!com_ESTADO
      txtdatos(4).text = com_cont!COM_TIPO_CTA
      txtdatos(5).text = com_cont!com_SIGNO_D
      txtdatos(6).text = com_cont!com_SIGNO_H
      txtdatos(7).text = com_cont!com_ACT_PAS
   End If
   com_cont.MoveNext
Loop
Posi_Ultimo = 0
NIVEL_ACT = 0
TreeView1.TabIndex = 0
barrita.Visible = False
If LK_TABLA = "BUSCAR" Then
 tcuenta.Visible = False
 tnombre.Visible = False
 etiqueta.Visible = False
 TreeView1.Height = 3135
End If
Exit Sub
sigue:
MsgBox Err.Description
'Resume Next
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
Cancel = True   ' Se cancela la operación
End Sub

Private Sub TreeView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  frmBuscacta.Hide
End If
If KeyAscii = 13 And (LK_TABLA = "BUSCAR" Or LK_TABLA = "SUB_TRA") Then
  If NIVEL_MAX = Val(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Tag) Then
    cmdcolocar_Click
  End If
End If

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
Dim wcuenta As String
Dim windex As String
If Trim(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Tag) = "RAIZ" Then
 NIVEL_ACT = 0
 GoTo FIN
End If
NIVEL_ACT = Val(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Tag)
COLOCA_CTA
FIN:
End Sub

Public Sub LLENA_CUENTAS()
PSCOM_CONT(0) = LK_CODCIA
com_cont.Requery
fila = 0
If com_cont.EOF Then Exit Sub
TreeView1.Nodes.Clear
Set nodX = TreeView1.Nodes.Add(, tvwChild, "TITULO", "Cuentas Principales")
TreeView1.Nodes.Item(1).Tag = "RAIZ"
TreeView1.Indentation = 400
Do Until com_cont.EOF
   fila = fila + 1
   If com_cont!COM_NIVEL = 1 Then
      Set nodX = TreeView1.Nodes.Add("TITULO", tvwChild, "A" + Left(com_cont!com_cuenta, wCOM_NIVEL(1)), Trim(com_cont!com_cuenta) + " " + Trim(com_cont!com_DESCRIPCION))
   ElseIf com_cont!COM_NIVEL = 2 Then
      Set nodX = TreeView1.Nodes.Add("A" + Left(com_cont!com_cuenta, wCOM_NIVEL(1)), tvwChild, "B" + Left(com_cont!com_cuenta, wCOM_NIVEL(2)), Trim(com_cont!com_cuenta) + "  " + Trim(com_cont!com_DESCRIPCION))
   ElseIf com_cont!COM_NIVEL = 3 Then
      Set nodX = TreeView1.Nodes.Add("B" + Left(com_cont!com_cuenta, wCOM_NIVEL(2)), tvwChild, "C" + Left(com_cont!com_cuenta, wCOM_NIVEL(3)), Trim(com_cont!com_cuenta) + "  " + Trim(com_cont!com_DESCRIPCION))
   ElseIf com_cont!COM_NIVEL = 4 Then
      Set nodX = TreeView1.Nodes.Add("C" + Left(com_cont!com_cuenta, wCOM_NIVEL(3)), tvwChild, "D" + Left(com_cont!com_cuenta, wCOM_NIVEL(4)), Trim(com_cont!com_cuenta) + "  " + Trim(com_cont!com_DESCRIPCION))
   ElseIf com_cont!COM_NIVEL = 5 Then
      Set nodX = TreeView1.Nodes.Add("D" + Left(com_cont!com_cuenta, wCOM_NIVEL(4)), tvwChild, "E" + Left(com_cont!com_cuenta, wCOM_NIVEL(5)), Trim(com_cont!com_cuenta) + "  " + Trim(com_cont!com_DESCRIPCION))
   ElseIf com_cont!COM_NIVEL = 6 Then
      Set nodX = TreeView1.Nodes.Add("E" + Left(com_cont!com_cuenta, wCOM_NIVEL(5)), tvwChild, "F" + Left(com_cont!com_cuenta, wCOM_NIVEL(6)), Trim(com_cont!com_cuenta) + "  " + Trim(com_cont!com_DESCRIPCION))
   End If
   TreeView1.Nodes.Item(fila + 1).Tag = Str(com_cont!COM_NIVEL)
 '  WS_CUENTA = com_cont!com_cuenta
 '  txtcuenta.text = com_cont!com_cuenta
 '  txtindice.text = ""
 '  If com_cont!COM_NIVEL <> 1 Then
 '     txtcuenta.text = Left(Trim(WS_CUENTA), wCOM_NIVEL(com_cont!COM_NIVEL - 1))
 '     txtindice.text = Right(Trim(WS_CUENTA), (wCOM_NIVEL(com_cont!COM_NIVEL) - wCOM_NIVEL(com_cont!COM_NIVEL - 1)))
 '  End If
   com_cont.MoveNext
Loop
Posi_Ultimo = 0
NIVEL_ACT = 0
Exit Sub

End Sub

Public Sub ACTUALIZA_CUENTAS(WCUENTAS As String, WTIPO As Integer, wsignod As Integer, wsignoh As Integer)
Dim wcuentas2 As String
pub_cadena = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? and COM_CUENTA > ?  ORDER BY COM_CUENTA"
Set PSCOM_MAYOR2 = CN.CreateQuery("", pub_cadena)
PSCOM_MAYOR2(0) = 0
PSCOM_MAYOR2(1) = 0
Set com_mayor2 = PSCOM_MAYOR2.OpenResultset(rdOpenKeyset, rdConcurValues)
PSCOM_MAYOR2(0) = LK_CODCIA
PSCOM_MAYOR2(1) = WCUENTAS
com_mayor2.Requery
If com_mayor2.EOF Then
  Exit Sub
End If
Do Until com_mayor2.EOF
   If com_mayor2!COM_NIVEL = 1 Then Exit Do
   com_mayor2.Edit
   com_mayor2!COM_TIPO_CTA = WTIPO
   com_mayor2!com_SIGNO_D = wsignod
   com_mayor2!com_SIGNO_H = wsignoh
   com_mayor2.Update
   com_mayor2.MoveNext
Loop

End Sub

Public Sub COPIA_ESTRU(WCUENTAS As String)
Dim wcuentas2 As String
Dim Wflag As String * 1
Dim wcuenta As Integer
Dim valor
Dim WULTIMO As Integer
valor = InputBox("Ingrese la Cuenta de Nivel Reemplasante, segun el nivel donde se encuentra  :", Pub_Titulo, "")
If valor = "" Then
  Screen.MousePointer = 0
  Exit Sub
End If
If Not IsNumeric(valor) Then
  MsgBox "Cuenta No Procede. ", 48, Pub_Titulo
  Exit Sub
End If
If NIVEL_ACT = 1 Then
  If Len(valor) <> Len(Left(WCUENTAS, wCOM_NIVEL(NIVEL_ACT))) Then
    MsgBox "Cuenta No es correcta. ", 48, Pub_Titulo
    Exit Sub
  End If
Else
  If Left(valor, wCOM_NIVEL(NIVEL_ACT - 1)) <> Left(WCUENTAS, wCOM_NIVEL(NIVEL_ACT - 1)) Then
    MsgBox "Cuenta No es correcta. ", 48, Pub_Titulo
    Exit Sub
  End If
End If

pub_cadena = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? and COM_CUENTA >= ?  ORDER BY COM_CUENTA"
Set PSCOM_MAYOR2 = CN.CreateQuery("", pub_cadena)
PSCOM_MAYOR2(0) = 0
PSCOM_MAYOR2(1) = 0
Set com_mayor2 = PSCOM_MAYOR2.OpenResultset(rdOpenKeyset, rdConcurValues)
PSCOM_MAYOR2(0) = LK_CODCIA
PSCOM_MAYOR2(1) = WCUENTAS
com_mayor2.Requery
wcuenta = 0
If com_mayor2.EOF Then
  MsgBox "Intente Nuevamente...", 48, Pub_Titulo
  Exit Sub
End If
'wfijo = Left(valor, wCOM_NIVEL(com_mayor2!COM_NIVEL - 1))
'windice = Mid(com_mayor2!com_cuenta, Len(Trim(wfijo)) + 1, Len(com_mayor2!com_cuenta))
SQ_OPER = 1
PUB_CUENTA = valor
LEER_COM_LLAVE
If Not com_llave.EOF Then
 MsgBox "Cuenta Existe en Plan Contable .", 48, Pub_Titulo
 Exit Sub
End If
PB.Min = 0
PB.Max = 3
PB.Value = 0
PB.Visible = True
DoEvents
Wflag = ""
WULTIMO = -1
PB.Value = PB.Value + 1
Do Until com_mayor2.EOF
   If NIVEL_ACT = com_mayor2!COM_NIVEL Then wcuenta = wcuenta + 1
   If wcuenta = 2 Then Exit Do
   If NIVEL_ACT > com_mayor2!COM_NIVEL Then Exit Do
   If Trim(Wflag) = "" Then
    Wflag = "A"
    WFIJO = Left(valor, wCOM_NIVEL(NIVEL_ACT))
    windice = ""
   Else
    WFIJO = Left(valor, wCOM_NIVEL(com_mayor2!COM_NIVEL - 1))
    windice = Mid(com_mayor2!com_cuenta, Len(Trim(WFIJO)) + 1, Len(com_mayor2!com_cuenta))
   End If
'  MsgBox wfijo & windice
'  GoTo dale
   com_llave.AddNew
   com_llave!com_cuenta = WFIJO + windice
   com_llave!COM_TIPO_CTA = com_mayor2!COM_TIPO_CTA
   com_llave!COM_CODCIA = LK_CODCIA
   com_llave!com_DESCRIPCION = com_mayor2!com_DESCRIPCION
   com_llave!COM_NIVEL = com_mayor2!COM_NIVEL
   com_llave!com_cuenta_sup = Left(Trim(com_llave!com_cuenta), wCOM_NIVEL(com_llave!COM_NIVEL - 1))
   com_llave!com_cuenta_AUTOM_D = ""
   com_llave!com_cuenta_AUTO_H = ""
   If com_mayor2!COM_NIVEL = NIVEL_MAX Then
       com_llave!com_FLAG_AFECTACION = "1"
   Else
       com_llave!com_FLAG_AFECTACION = "0"
   End If
   com_llave!com_ESTADO = ""
   com_llave!COM_DEB_ANO = 0
   com_llave!COM_HAB_ANO = 0
   com_llave!COM_DEB_MES = 0
   com_llave!COM_HAB_MES = 0
   com_llave!com_SIGNO_D = com_mayor2!com_SIGNO_D
   com_llave!com_SIGNO_H = com_mayor2!com_SIGNO_H
   com_llave!com_ACT_PAS = 0
   com_llave!COM_TIPO_CTA = com_mayor2!COM_TIPO_CTA
   com_llave.Update
   If com_mayor2!COM_NIVEL = 1 Then
     TreeView1.Nodes.Add "TITULO", tvwChild, "A", "Nueva Cuenta..."
   ElseIf com_mayor2!COM_NIVEL = 2 Then
    TreeView1.Nodes.Add TreeView1.Nodes.Item("A" + Left(WFIJO + windice, wCOM_NIVEL(com_mayor2!COM_NIVEL - 1))).Key, tvwChild, "B", "Nueva Cuenta..."
   ElseIf com_mayor2!COM_NIVEL = 3 Then
    TreeView1.Nodes.Add TreeView1.Nodes.Item("B" + Left(WFIJO + windice, wCOM_NIVEL(com_mayor2!COM_NIVEL - 1))).Key, tvwChild, "C", "Nueva Cuenta..."
   ElseIf com_mayor2!COM_NIVEL = 4 Then
    TreeView1.Nodes.Add TreeView1.Nodes.Item("C" + Left(WFIJO + windice, wCOM_NIVEL(com_mayor2!COM_NIVEL - 1))).Key, tvwChild, "D", "Nueva Cuenta..."
   ElseIf com_mayor2!COM_NIVEL = 5 Then
    TreeView1.Nodes.Add TreeView1.Nodes.Item("D" + Left(WFIJO + windice, wCOM_NIVEL(com_mayor2!COM_NIVEL - 1))).Key, tvwChild, "E", "Nueva Cuenta..."
   ElseIf com_mayor2!COM_NIVEL = 6 Then
    TreeView1.Nodes.Add TreeView1.Nodes.Item("E" + Left(WFIJO + windice, wCOM_NIVEL(com_mayor2!COM_NIVEL - 1))).Key, tvwChild, "F", "Nueva Cuenta..."
   End If
   Set TreeView1.SelectedItem = TreeView1.Nodes.Item(TreeView1.Nodes.Count)
   TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).EnsureVisible
   TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).text = Trim(WFIJO + windice) + "  " + Trim(com_mayor2!com_DESCRIPCION)
   TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key = TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key + Trim(WFIJO + windice)
   TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Tag = com_mayor2!COM_NIVEL
   If WULTIMO = -1 Then
     WULTIMO = TreeView1.SelectedItem.Index
   End If
   com_mayor2.MoveNext
Loop
PB.Value = PB.Value + 1
TreeView1.Nodes.Item(1).Sorted = True
If WULTIMO <> -1 Then
 Set TreeView1.SelectedItem = TreeView1.Nodes.Item(WULTIMO)
 TreeView1_NodeClick TreeView1.Nodes.Item(TreeView1.SelectedItem.Index)
End If
PB.Value = PB.Value + 1
PB.Visible = False
MsgBox "Proceso de Copiado Terminado.", 48, Pub_Titulo
End Sub

Public Sub COLOCA_CTA()
Dim LOC_CUENTA As String
Dim PSCOM_MAYOR2 As rdoQuery
Dim com_mayor2 As rdoResultset
Dim wcuenta As String
Dim wnivel As Integer
Dim indice As Integer
Dim nivel_sigue As Integer
If TreeView1.Nodes.Count = 0 Then
 Exit Sub
End If
If Nulo_Valors(par_llave!PAR_CONTA_C) <> "A" And Trim(LK_TABLA) = "CLIENTE" Then
 GoTo VERIFICA
End If
If Nulo_Valors(par_llave!PAR_CONTA_P) <> "A" And Trim(LK_TABLA) = "PROVEEDOR" Then
 GoTo VERIFICA
End If
If Trim(LK_TABLA) = "PROVEEDOR2" Or Trim(LK_TABLA) = "CLIENTES2" Then
  GoTo VERIFICA
End If
If Trim(LK_TABLA) = "ARTI1" Or Trim(LK_TABLA) = "ARTI2" Or Trim(LK_TABLA) = "ARTI3" Or Trim(LK_TABLA) = "ARTI4" Then
 GoTo VERIFICA
End If
If Trim(LK_TABLA) = "BUSCAR" Then
 GoTo VERIFICA
End If
If Trim(LK_TABLA) = "SUB_TRA" Then
 GoTo VERIFICA
End If

If NIVEL_ACT <> NIVEL_MAX - 1 Then
        tcuenta.text = ""
        tnombre.text = ""
        Exit Sub
End If
wcuenta = Mid(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key, 2, Len(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key))
If wcuenta = "ITULO" Then
 MsgBox "No procede .", 48, Pub_Titulo
 Exit Sub
Else
 LOC_CUENTA = Left(wcuenta, wCOM_NIVEL(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Tag))
End If
'nivel_sigue = wCOM_NIVEL(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Tag) + 1
Posi_Ultimo = TreeView1.SelectedItem.Index
cade = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_CUENTA >= ? and COM_CUENTA < ? ORDER BY COM_CUENTA"
Set PSCOM_MAYOR2 = CN.CreateQuery("", cade)
PSCOM_MAYOR2(0) = 0
PSCOM_MAYOR2(1) = 0
PSCOM_MAYOR2(2) = 0
Set com_mayor2 = PSCOM_MAYOR2.OpenResultset(rdOpenKeyset, rdConcurValues)

wnivel = NIVEL_ACT + 1
If wnivel = 1 Then
 MsgBox "No Procede.", 48, Pub_Titulo
 Exit Sub
End If
If wnivel > NIVEL_MAX Then
 MsgBox "No Procede, en el Ultimo Nivel. ", 48, Pub_Titulo
 Exit Sub
End If
If wnivel = 1 Then
  VAR1 = wCOM_NIVEL(wnivel)
Else
  VAR1 = wCOM_NIVEL(wnivel) - Len(LOC_CUENTA)
End If

PSCOM_MAYOR2.rdoParameters(0) = LK_CODCIA
PSCOM_MAYOR2.rdoParameters(1) = LOC_CUENTA & "01"
PSCOM_MAYOR2.rdoParameters(2) = LOC_CUENTA & "99" 'Val(LOC_CUENTA) + 1
com_mayor2.Requery
wcuenta = "X"
xcuenta = 0
If Not com_mayor2.EOF Then
  com_mayor2.MoveLast
  wcuenta = Trim(com_mayor2!com_cuenta)
  WDESCRIP = Trim(com_mayor2!com_DESCRIPCION)
' Do Until com_mayor2.EOF
'   If wnivel > com_mayor2!com_NIVEL Then Exit Do
'   If wnivel = com_mayor2!com_NIVEL Then
'     wcuenta = Trim(com_mayor2!com_cuenta)
'     WDESCRIP = Trim(com_mayor2!com_descripcion)
'   End If
'   com_mayor2.MoveNext
' Loop
End If

If wcuenta = "X" Then
 indice = 1
Else
 indice = Val(Right(wcuenta, VAR1)) + 1
End If
 If VAR1 = 1 Then
  wcuenta = LOC_CUENTA & Format(indice, "0")
 ElseIf VAR1 = 2 Then
  wcuenta = LOC_CUENTA & Format(indice, "00")
 ElseIf VAR1 = 3 Then
  wcuenta = LOC_CUENTA & Format(indice, "000")
 ElseIf VAR1 = 4 Then
  wcuenta = LOC_CUENTA & Format(indice, "0000")
 ElseIf VAR1 = 5 Then
  wcuenta = LOC_CUENTA & Format(indice, "00000")
 ElseIf VAR1 = 6 Then
  wcuenta = LOC_CUENTA & Format(indice, "000000")
 ElseIf VAR1 = 7 Then
  wcuenta = LOC_CUENTA & Format(indice, "0000000")
 ElseIf VAR1 = 8 Then
  wcuenta = LOC_CUENTA & Format(indice, "00000000")
 End If
frmBuscacta.tcuenta.text = wcuenta
If Trim(LK_TABLA) = "CLIENTE" Or Trim(LK_TABLA) = "PROVEEDOR" Or Trim(LK_TABLA) = "PROVEEDOR2" Then
 frmBuscacta.tnombre.text = (frmCLI.txtesposo.text)
ElseIf Trim(LK_TABLA) = "ARTI1" Or Trim(LK_TABLA) = "ARTI2" Or Trim(LK_TABLA) = "ARTI3" Or Trim(LK_TABLA) = "ARTI4" Then
 frmBuscacta.tnombre.text = Trim(frmARTI.txtnombre.text)
ElseIf Trim(LK_TABLA) = "BANCO" Then
 frmBuscacta.tnombre.text = Trim(frmBancos.TXTBAN(2).text)
Else
 frmBuscacta.tnombre.text = WDESCRIP
End If

Exit Sub
VERIFICA:
  If NIVEL_ACT <> NIVEL_MAX Then
       tcuenta.text = ""
       tnombre.text = ""
       Exit Sub
  End If
  wcuenta = Mid(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key, 2, Len(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).Key))
  frmBuscacta.tcuenta.text = wcuenta
  wcuenta = Mid(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).text, wCOM_NIVEL(NIVEL_ACT) + 1, Len(TreeView1.Nodes.Item(TreeView1.SelectedItem.Index).text))
  frmBuscacta.tnombre.text = Trim(wcuenta)
End Sub
