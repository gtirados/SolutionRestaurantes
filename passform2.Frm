VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PASSFORM 
   BackColor       =   &H00FAEFDA&
   BorderStyle     =   0  'None
   Caption         =   "ADMIN - Ingresos de Usuario"
   ClientHeight    =   5490
   ClientLeft      =   1005
   ClientTop       =   1995
   ClientWidth     =   4830
   ForeColor       =   &H00FFFF00&
   Icon            =   "passform2.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   1  'Arrow
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5490
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton SALIR 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Salir del Sistema"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1380
      Picture         =   "passform2.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Frame F2 
      BackColor       =   &H00FAEFDA&
      Height          =   3855
      Left            =   120
      TabIndex        =   5
      Top             =   510
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton cmdIngresar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Ok."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   3240
         Picture         =   "passform2.frx":0CB8
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2265
         Width           =   1095
      End
      Begin VB.CheckBox cheguarda 
         BackColor       =   &H00FAEFDA&
         Caption         =   "&Guardar Password "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008B4914&
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text2 
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
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin MSMask.MaskEdBox fecha_proceso 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1545
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Compañia por Defecto :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008B4914&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008B4914&
         Height          =   255
         Left            =   135
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha del Sistema :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008B4914&
         Height          =   435
         Left            =   135
         TabIndex        =   11
         Top             =   1920
         Width           =   1065
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008B4914&
         Height          =   315
         Left            =   150
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.Label LBLPAS 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008B4914&
         Height          =   315
         Left            =   135
         TabIndex        =   8
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   7
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Momento 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Un Momento  !!!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008B4914&
         Height          =   240
         Left            =   1410
         TabIndex        =   6
         Top             =   3600
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label Label6 
         BackColor       =   &H00EAC793&
         BackStyle       =   0  'Transparent
         Caption         =   "Nota: Para ingresar con otro USUARIO pulse [Esc]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   135
         TabIndex        =   13
         Top             =   2520
         Width           =   1935
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   2520
   End
   Begin VB.Label lblbarraos 
      Alignment       =   2  'Center
      BackColor       =   &H008B4914&
      Caption         =   "Solution for Business"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   5160
      Width           =   4815
   End
   Begin VB.Label empresa 
      Alignment       =   2  'Center
      BackColor       =   &H008B4914&
      Caption         =   "Empresa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblempresa 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1215
      TabIndex        =   12
      Top             =   120
      Width           =   3480
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   1185
   End
End
Attribute VB_Name = "PASSFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, J As Integer
Dim PASSW  As String * 10
Dim VECES As Integer
Dim LOC_FLAG_OJO As String
Public Function WRITE1_READ2(WMODO As Integer, Optional wWRITE) As String
Dim RUTA As String
On Error GoTo SALE
RUTA = "C:\" & "CODUSU$.txt"

Select Case WMODO
Case 1
    Open RUTA For Output As #1
    Print #1, wWRITE
    Close #1
Case 2
    Dim WLINEA
    Open RUTA For Input As #1
    Do While Not EOF(1)
        Line Input #1, WLINEA
    Loop
    Close #1
    WRITE1_READ2 = WLINEA
Case Else
   MsgBox "wMODO 1 o 2 ...", 48, Pub_Titulo
End Select

Exit Function


SALE:
 If Err.Number = 53 Then
    WRITE1_READ2 = ""
 Else
    MsgBox Err.Number & Err.Description, 48, Pub_Titulo
 End If
 
 Close #1
 

End Function



Private Sub cmdIngresar_Click()
Dim WRES
Dim mensa As String
Dim pub_mensajeText
Dim PB
Dim wfecha_bloq As Date
Dim wserie As Currency
PB = Chr(10) & Chr(13) & Chr(10) & Chr(13)
On Error GoTo SALE

If IsDate(fecha_proceso.Text) = False Then
   MsgBox "Fecha Invalida ...", 48, Pub_Titulo
   fecha_proceso.SetFocus
   fecha_proceso.SelStart = 0
   fecha_proceso.SelLength = Len(fecha_proceso.Text)
   GoTo fin
End If
'If LK_CODUSU = "ADMIN" Then GoTo pasa_todo
If CDate(fecha_proceso.Text) < 1900 Then
   MsgBox "Fecha NO Procede ...", 48, Pub_Titulo
   fecha_proceso.SetFocus
   fecha_proceso.SelStart = 0
   fecha_proceso.SelLength = Len(fecha_proceso.Text)
   GoTo fin
End If
PUB_CODCIA = LK_CODCIA
PUB_CAL_INI = fecha_proceso.Text
PUB_CAL_FIN = fecha_proceso.Text
SQ_OPER = 1
LEER_CAL_LLAVE 1
If cal_llave.EOF Then
   MsgBox "Fecha de Proceso Errada ...", 48, Pub_Titulo
   fecha_proceso.SetFocus
   fecha_proceso.SelStart = 0
   fecha_proceso.SelLength = Len(fecha_proceso.Text)
   GoTo fin
End If
If cal_llave!CAL_INDICE <> 1 Then
   MsgBox "Fecha no Activada ... el Calendario del Sistema No Actualizo !!!", 48, Pub_Titulo
   fecha_proceso.SetFocus
   fecha_proceso.SelStart = 0
   fecha_proceso.SelLength = Len(fecha_proceso.Text)
   If LK_CODUSU = "ADMIN" Then
     MsgBox "Usuario Administrador, corregir el calendario del Sistema .", 48
     GoTo SALTA_TODO
   End If
   GoTo fin
End If
  ' VER SI LA FECHA NO ES DEL DIA
SALTA_TODO:
  pub_mensaje = "¡La fecha de la compañia es diferente que el computador!"
  pub_mensaje = pub_mensaje & PB & "Fecha de la Compañia       : " & Format(fecha_proceso, "dd mmm yyyy")
  pub_mensaje = pub_mensaje & Chr(13) & "Fecha de del Computador : " & Format(Date, "dd mmm yyyy")
  pub_mensaje = pub_mensaje & PB & "¿Desea ingresar al sistema..?"
  If CDate(fecha_proceso.Text) <> Date Then
    Pub_Respuesta = MsgBox(pub_mensaje, vbOKCancel + vbInformation, Pub_Titulo)
    If Pub_Respuesta = vbCancel Then
       fecha_proceso.SetFocus
       fecha_proceso.SelStart = 0
       fecha_proceso.SelLength = Len(fecha_proceso.Text)
       Exit Sub
    End If
  End If
pasa_todo:
  PASSFORM.Timer1.Enabled = False
  LK_FECHA_DIA = fecha_proceso.Text
  'LK_FECHA_AYER = Nulo_Valor0(gen!GEN_FECHA_AYER)
  If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Then
   PUB_CLAVE = Trim(Text1.Text)
  Else
    usu.Requery
    Do Until usu.EOF
      If Trim(usu!USU_KEY) = "SUPERVISOR" Then
        PUB_CLAVE = Trim(usu!USU_CLAVE)
        Exit Do
      End If
      usu.MoveNext
    Loop
  End If
  
  Dim wGUARDA As String
  wGUARDA = Trim(Text2.Text) & CStr(cheguarda.Value)
  DoEvents
  Call WRITE1_READ2(1, wGUARDA)
  momento.Visible = True
  DoEvents
  momento.Caption = " O K .  ==  >>  " & momento.Caption
  PASSFORM.Hide
  MDIForm1.Show
Exit Sub
SALE:
 MsgBox "Reiciar Sistema ..Intente Nuevamente...", 48, Pub_Titulo
  End
 
fin:
End Sub

Private Sub F1_Click()

End Sub

Private Sub fecha_proceso_Change()
If IsDate(fecha_proceso.Text) Then
   momento.Visible = True
   momento.Caption = Format(fecha_proceso, "dddd, d mmmm yyyy")
Else
   momento.Visible = False
End If

End Sub

Private Sub Form_Activate()
LOC_FLAG_OJO = ""
VECES = 0
Screen.MousePointer = 11
DoEvents
'PASSFORM.F1.Visible = True
DoEvents
PASSFORM.F2.Visible = True
DoEvents
PASSFORM.momento.Caption = "Cargando Usuarios ... " + PASSFORM.momento.Caption
DoEvents
PASSFORM.momento.Visible = True
DoEvents
Text2.Enabled = True
DoEvents
PASSFORM.momento.Visible = False
Dim WLON, WTODO
WTODO = WRITE1_READ2(2)
If WTODO = "" Then
   Screen.MousePointer = 0
   Exit Sub
End If
Text2.Text = Left(WTODO, Len(WTODO) - 1)
If Len(Text2.Text) = 0 Then
   Screen.MousePointer = 0
   Exit Sub
End If
Dim wCHEK As String
If Trim(Nulo_Valors(GEN!gen_bloqueo)) = "A" Then
  Text2.Text = ""
  Text2.Enabled = True
  Text2.SetFocus
  GoTo SALTA
End If
wCHEK = Right(WTODO, 1)
cheguarda.Value = val(wCHEK)
If wCHEK = "0" Then
    Text2_KeyPress 13
ElseIf wCHEK = "1" Then
    Text2_KeyPress 13
    Text1.Text = Trim(PASSW)
    Text1_KeyPress 13
End If
SALTA:

Screen.MousePointer = 0
End Sub
Private Sub fecha_proceso_KeyPress(KeyAscii As Integer)
Dim WRES
Dim mensa As String
If KeyAscii = 27 Then
   VECES = 0
   cheguarda.Enabled = False
   cmdIngresar.Enabled = False
   Timer1.Enabled = False
   LBLPAS.Visible = True
   Text1.Text = ""
   Text1.Enabled = True
   fecha_proceso.Enabled = False
   Text1.SetFocus
  Exit Sub
End If
    
If KeyAscii <> 13 Then
     GoTo fin
End If
cmdIngresar_Click
fin:
End Sub

Private Sub Form_Load()

LOC_FLAG_OJO = ""
PS_GEN(0) = 0
GEN.Requery
If GEN.EOF Then
 MsgBox "La Tabla General no Existe Datos"
Else
 lblempresa.Caption = Trim(GEN!GEN_NOMBRE)
End If
End Sub

Private Sub salir_Click()
        Screen.MousePointer = 11
        DoEvents
        PASSFORM.Hide
        MDIForm1.WindowState = 1
        MDIForm1.Caption = ""
        EN.Close
        CN.Close
        Pub_ConnAdo.Close
        Screen.MousePointer = 0
        End
End Sub

Private Sub Text1_GotFocus()
Timer1.Enabled = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim whasta As Integer
Dim wGUARDA As String
ALFA_NUMERICO KeyAscii
whasta = 4
If KeyAscii = 13 Or KeyAscii = 27 Then
Else
   GoTo fin
End If

If KeyAscii = 27 Then
   VECES = 0
  Timer1.Enabled = False
  Label3.Caption = ""
  LBLPAS.Visible = True
  Text1.Text = ""
  Text2.Text = ""
  Text2.Enabled = True
  Text2.SetFocus
  Exit Sub
End If
VECES = VECES + 1
If VECES > whasta Then
    MsgBox " PERSONA NO GRATA AL SISTEMA ......!!!!!", vbCritical, Pub_Titulo
    usu.Edit
    usu!usu_flag_ojo = "A"
    usu.Update
    End
End If

If UCase(Trim(Text1.Text)) <> UCase(Trim(PASSW)) Then
   MsgBox "Password Incorrecto", vbCritical + vbDefaultButton2, Pub_Titulo
   If Text1.Enabled Then Text1.SetFocus
   Text1.Text = ""
   GoTo fin
End If

If KeyAscii = 13 Then
If VECES = 4 Then
   pub_mensaje = "Persona  NO  grata al  SISTEMA ..."
   MsgBox pub_mensaje, 48, Pub_Titulo
    usu.Edit
    usu!usu_flag_ojo = "A"
    usu.Update
   End
End If
End If
PS_GEN(0) = 0
GEN.Requery
Text1.Enabled = False
fecha_proceso.Enabled = True
SQ_OPER = 1
PUB_CODCIA = usu!usu_CODCIA
LEER_PAR_LLAVE
fecha_proceso.Text = Format(par_llave!PAR_FECHA_DIA, "dd/mm/yyyy")
cheguarda.Enabled = True
cmdIngresar.Enabled = True
fecha_proceso.Mask = "##/##/####"
fecha_proceso.SetFocus
fecha_proceso.SelStart = 0
fecha_proceso.SelLength = Len(fecha_proceso.Text)
If par_llave!PAR_EMPRESA = "PLA" Then
 ' If XLL Is Nothing Then
 '    Set XLL = CreateObject("Excel.Application")
 ' End If
  'DoEvents
'  XLL.Workbooks.Open Trim(PUB_RUTA_OTRO) & "TASAS.xls", 0, True, 4, PUB_CLAVE, PUB_CLAVE
End If

If LOC_FLAG_OJO = "A" Then
  MsgBox "Hubo un intento de acceso incorrecto con su Usuario.!!!  Seria conveniente que cambie su Clave.", 48, Pub_Titulo
  usu.Edit
  usu!usu_flag_ojo = " "
  usu.Update
End If
momento.Caption = Trim(par_llave!PAR_NOMBRE) & " Fecha: " & Format(par_llave!PAR_FECHA_DIA, "dd/mm/yy")
momento.Visible = True


fin:
End Sub

Private Sub Text1_LostFocus()
Timer1.Enabled = False
LBLPAS.Visible = True
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim car As String, Longt As Integer
Dim RES As String * 1
Dim WSrep As Integer
WSrep = 4  'Repite cuentas veces tiene de opcion
car = Chr$(KeyAscii)
car = UCase$(Chr$(KeyAscii))
KeyAscii = Asc(car)
If car < "A" Or car > "Z" Then
  If KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
        KeyAscii = 0
'        Beep
    End If
End If
If KeyAscii = 13 Then
   valor = Trim(Text2.Text)
   RES = "N"
   usu.Requery
   usu.MoveFirst
   Do Until usu.EOF
     If Trim(usu!USU_KEY) = valor Then
        RES = "S"
        Exit Do
     End If
     usu.MoveNext
   Loop
   If RES = "N" Then
       pub_mensaje = "Usuario Incorrecto"
       Pub_Titulo = "GT Software S.A.C. - Solution"
       Text2.Text = ""
       If VECES < WSrep - 1 Then
          MsgBox pub_mensaje, 48, "Solution"
       End If
    Else
       GoTo PASO_TODO
    End If
End If
If KeyAscii = 13 Then
   VECES = VECES + 1
   If VECES = WSrep Then
      MsgBox "Persona  NO   grata al SISTEMA  .... ", 48, "Admin-Soft"
     End 'FIN DE PROGR.
   End If
End If
Exit Sub
PASO_TODO:
       i = 1
       Do Until i > 10
          lk_GRUPOS(i) = usu(3 + i)
          i = i + 1
       Loop
       i = 1
       Do Until i > 10
          lk_CODTRAS(i) = usu(13 + i)
          i = i + 1
       Loop
       
       'Correcto
       cheguarda.Enabled = False
       cmdIngresar.Enabled = False
       USUARIO = Text2.Text
       Label3.Caption = usu!USU_NOMBRE
       PASSW = usu!USU_CLAVE
       Codi_Grupo = usu!USU_GRUPO1
       LK_CODUSU = Trim(usu!USU_KEY)
       LK_CODCIA = usu!usu_CODCIA
       LK_FLAG_LIMITE = Nulo_Valors(usu!USU_LIMITE)
       LK_FAC_IMP = Nulo_Valors(usu!usu_fac_imp)
       LK_USU_STOCK = Nulo_Valors(usu!usu_stock)
       LK_USU_CUENTA = Nulo_Valors(usu!usu_cuenta)
       LK_USU_IMPRIME = Nulo_Valors(usu!usu_imprime)
       LK_USU_PASSWORD = Nulo_Valors(usu!usu_password)
       LOC_FLAG_OJO = Nulo_Valors(usu!usu_flag_ojo)
       LK_DEVICE_FBG = Nulo_Valors(usu!USU_DEVICE_FBG)
       Text2.Enabled = False
       Timer1.Enabled = True
       Text1.Enabled = True
       Text1.SetFocus
       
       
    'End If



End Sub


Private Sub Timer1_Timer()
LBLPAS.Visible = Not LBLPAS.Visible
End Sub

Public Function INGRESO_PASS(TEXTO) As String
cade = " ABCDEFGHIJKLMNOPQSRTVWXYZ0123456789"
SIMB = "ò~!@#$%^&*¤¶§éû¬½¼©¥+ãÃÓ¦¢ƒ×Ø£øÜïëÇ"
TOT = Len(TEXTO)
For i = 1 To TOT
 tem = InStr(1, cade, Mid(UCase(TEXTO), i, 1))
 CARAC = Mid(SIMB, tem, 1)
 CADENA = CADENA + CARAC
Next i
INGRESO_PASS = CADENA
End Function

Public Function SALIDA_PASS(TEXTO) As String
cade = " ABCDEFGHIJKLMNOPQSRTVWXYZ0123456789"
SIMB = "ò~!@#$%^&*¤¶§éû¬½¼©¥+ãÃÓ¦¢ƒ×Ø£øÜïëÇ"
TOT = Len(TEXTO)
For i = 1 To TOT
 tem = InStr(1, SIMB, Mid(TEXTO, i, 1))
 CARAC = Mid(cade, tem, 1)
 CADENA = CADENA + CARAC
Next i
SALIDA_PASS = CADENA
End Function
Public Sub ALFA_NUMERICO(Optional tecla)
'CONVIERTE TODA A MAYUSCULAS LETRAS
Dim car As String, Longt As Integer
car = Chr$(tecla)
car = UCase$(Chr$(tecla))
tecla = Asc(car)
If car < "A" Or car > "Z" Then
If car < "0" Or car > "9" Then
    If tecla <> 8 And tecla <> 13 And tecla <> 27 And tecla <> 32 Then
        tecla = 0
'        Beep
    End If
End If
End If
End Sub


