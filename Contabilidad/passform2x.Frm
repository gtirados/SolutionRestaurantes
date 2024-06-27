VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PASSFORM 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "ADMIN - Ingresos de Usuario"
   ClientHeight    =   3540
   ClientLeft      =   1005
   ClientTop       =   1995
   ClientWidth     =   8775
   ForeColor       =   &H00FFFF00&
   Icon            =   "passform2.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   1  'Arrow
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3540
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame F2 
      Height          =   3255
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdIngresar 
         Caption         =   "&Aceptar"
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
         Height          =   492
         Left            =   3480
         TabIndex        =   4
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CheckBox cheguarda 
         Caption         =   "&Guardar Clave"
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
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox Text2 
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
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton SALIR 
         Caption         =   "&Salir"
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
         Left            =   3360
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1575
      End
      Begin MSMask.MaskEdBox fecha_proceso 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   840
         Width           =   1545
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   435
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1425
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   435
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   825
      End
      Begin VB.Label LBLPAS 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clave :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label Label3 
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
         Height          =   360
         Left            =   3240
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Momento 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   " Un Momento  !!!"
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
         Height          =   240
         Left            =   720
         TabIndex        =   7
         Top             =   2880
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label Label6 
         Caption         =   "[ESC]= Otro Usuario"
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
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2640
         Width           =   2175
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   2400
   End
   Begin VB.Label lblempresa 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "ADMIN - SOFT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   240
      TabIndex        =   13
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label F1 
      BackColor       =   &H00000080&
      ForeColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   3255
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
      TabIndex        =   11
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
RUTA = "C:\CODUSU$.txt"

Select Case WMODO
Case 1
    Open RUTA For Output As #1
    Print #1, wWRITE
    Close #1
Case 2
    Dim wLINEA
    Open RUTA For Input As #1
    Do While Not EOF(1)
        Line Input #1, wLINEA
    Loop
    Close #1
    WRITE1_READ2 = wLINEA
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
Dim wres
Dim mensa As String
Dim pub_mensajeText
Dim pb
Dim wfecha_bloq As Date
pb = Chr(10) & Chr(13) & Chr(10) & Chr(13)
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
GoTo SALTA_TODO
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
  pub_mensaje = "<< OJO >> Fecha de esta Empresa NO es del Dia ..."
  pub_mensaje = pub_mensaje & pb & "Fecha del Sistema : " & Format(fecha_proceso, "dddd, d mmmm yyyy") & "            "
  pub_mensaje = pub_mensaje & pb & "Fecha de Hoy : " & Format(Date, "dddd, d mmmm yyyy") & "            "
  pub_mensaje = pub_mensaje & pb & "Desea Continuar ?."
  'If CDate(fecha_proceso.Text) <> Date Then
  '  Pub_Respuesta = MsgBox(pub_mensaje, 292, Pub_Titulo)
  '  If Pub_Respuesta = vbNo Then
  '     fecha_proceso.SetFocus
  '     fecha_proceso.SelStart = 0
  '     fecha_proceso.SelLength = Len(fecha_proceso.Text)
  '     Exit Sub
  '  End If
  'End If
pasa_todo:
  PASSFORM.Timer1.Enabled = False
  LK_FECHA_DIA = fecha_proceso.Text
  LK_FECHA_DIA = Format(Date, "dd/mm/yyyy")
  
  'LK_FECHA_AYER = Nulo_Valor0(gen!GEN_FECHA_AYER)
  If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Then
   PUB_CLAVE = Trim(Text1.Text)
  Else
    usu.Requery
    Do Until usu.EOF
      If Trim(usu!usu_key) = "SUPERVISOR" Then
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
PASSFORM.F1.Visible = True
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
wCHEK = Right(WTODO, 1)
cheguarda.Value = Val(wCHEK)
If wCHEK = "0" Then
    Text2_KeyPress 13
ElseIf wCHEK = "1" Then
    Text2_KeyPress 13
    Text1.Text = Trim(PASSW)
    Text1_KeyPress 13
End If


Screen.MousePointer = 0
End Sub
Private Sub fecha_proceso_KeyPress(KeyAscii As Integer)
Dim wres
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
 lblempresa.Caption = Trim(GEN!gen_nombre)
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
   Text1.SetFocus
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
If LOC_FLAG_OJO = "A" Then
  MsgBox "Habido un acceso incorecto en su clave de Ingreso.!!!  Seria conveniente que cambie su Password.", 48, Pub_Titulo
  usu.Edit
  usu!usu_flag_ojo = " "
  usu.Update
End If
momento.Caption = "Empresa : " & Trim(par_llave!PAR_NOMBRE) & "  - " & Format(par_llave!PAR_FECHA_DIA, "dddd, d mmmm yyyy")
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
     If Trim(usu!usu_key) = valor Then
        RES = "S"
        Exit Do
     End If
     usu.MoveNext
   Loop
   If RES = "N" Then
       pub_mensaje = "Usuario Incorrecto"
       Pub_Titulo = "SOLUTIN"
       Text2.Text = ""
       If VECES < WSrep - 1 Then
          MsgBox pub_mensaje, 48, "Admin-Soft"
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
       usuario = Text2.Text
       Label3.Caption = usu!USU_NOMBRE
       PASSW = usu!USU_CLAVE
       Codi_Grupo = usu!USU_GRUPO1
       LK_CODUSU = Trim(usu!usu_key)
       LK_CODCIA = usu!usu_CODCIA
       LK_FLAG_LIMITE = Nulo_Valors(usu!USU_LIMITE)
       LK_FAC_IMP = Nulo_Valors(usu!usu_fac_imp)
       LK_USU_STOCK = Nulo_Valors(usu!usu_stock)
       LOC_FLAG_OJO = Nulo_Valors(usu!usu_flag_ojo)
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
tot = Len(TEXTO)
For i = 1 To tot
 tem = InStr(1, cade, Mid(UCase(TEXTO), i, 1))
 CARAC = Mid(SIMB, tem, 1)
 CADENA = CADENA + CARAC
Next i
INGRESO_PASS = CADENA
End Function

Public Function SALIDA_PASS(TEXTO) As String
cade = " ABCDEFGHIJKLMNOPQSRTVWXYZ0123456789"
SIMB = "ò~!@#$%^&*¤¶§éû¬½¼©¥+ãÃÓ¦¢ƒ×Ø£øÜïëÇ"
tot = Len(TEXTO)
For i = 1 To tot
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

