VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "Comctl32.ocx"
Begin VB.Form Splash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1905
   ClientLeft      =   1935
   ClientTop       =   1560
   ClientWidth     =   7890
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Splash.frx":0000
   ScaleHeight     =   1905
   ScaleWidth      =   7890
   WhatsThisHelp   =   -1  'True
   Begin ComctlLib.ProgressBar rctStatusBar 
      Height          =   165
      Left            =   1920
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   291
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Label Empresa 
      BackStyle       =   0  'Transparent
      Caption         =   "Transferencia de Datos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   6240
      TabIndex        =   2
      Top             =   4080
      Width           =   465
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando Módulos ...."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3360
      TabIndex        =   0
      Top             =   4275
      Width           =   1800
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iStatusBarWidth As Integer

Private Sub Form_Click()
' Unload Me
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo SALE
If KeyCode = 120 Then
    Splash.WindowState = 1
    Splash.Caption = "Proceso Abortado!!! . . ."
    Screen.MousePointer = 0
    EN.Close
    CN.Close
    Screen.MousePointer = 0
    End
End If
Exit Sub
SALE:
End
End Sub

Private Sub Form_Load()
Dim success%
Dim pb
pb = Chr(10) & Chr(13) & Chr(10) & Chr(13)
'On Error GoTo SALE
Screen.MousePointer = 11
If App.PrevInstance Then
  pub_mensaje = App.Path & " " & "SOLUTIN"
  pub_mensaje = pub_mensaje & pb & "Posiblemente la Aplicación este cargada o no ha sido cerrada Correctamente "
  pub_mensaje = pub_mensaje & pb & "Debe Cerrar todos los Programas e Iniciar la seccion como Usuario Distinto ..."
  MsgBox pub_mensaje, vbCritical, "SOLUTIN"
  Screen.MousePointer = 0
  End
End If

Pub_Titulo = "SOLUTIN"
LK_CODCIA = ""
LK_CODUSU = ""
If Nulo_Valor0(PUB_FLAG) = 0 Then
  Splash.rctStatusBar.max = 4300
  Splash.rctStatusBar.Min = 0
  Splash.rctStatusBar.Value = 0
  Splash.rctStatusBar.Visible = True
  DoEvents
  Splash.Show
  'success% = SetWindowPos(Splash.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
  DoEvents
  CONEXION_GEN
End If
PUB_FLAG = 0
DoEvents
'success% = SetWindowPos(Splash.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
DoEvents
Load MDIForm1
Screen.MousePointer = 0
Exit Sub
SALE:
 Screen.MousePointer = 0
 MsgBox Err.Description, 48, "pub_titulo"
End

End Sub

