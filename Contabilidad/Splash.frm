VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Splash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FAEDE9&
   BorderStyle     =   0  'None
   ClientHeight    =   3540
   ClientLeft      =   3165
   ClientTop       =   2235
   ClientWidth     =   5265
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
   ScaleHeight     =   3540
   ScaleWidth      =   5265
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.ProgressBar rctStatusBar 
      Height          =   165
      Left            =   120
      TabIndex        =   2
      Top             =   3030
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Autorizado a:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   165
      TabIndex        =   8
      Top             =   1605
      Width           =   1080
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H009C3000&
      BorderWidth     =   2
      Height          =   3525
      Left            =   15
      Top             =   15
      Width           =   5250
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gestión Contable"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2865
      TabIndex        =   7
      Top             =   750
      Width           =   2250
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "for Business"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1290
      TabIndex        =   6
      Top             =   495
      Width           =   1005
   End
   Begin VB.Label lblanexo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Solution"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   165
      TabIndex        =   3
      Top             =   135
      Width           =   1275
   End
   Begin VB.Label lblEmpresa 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando Modulos . . ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   225
      TabIndex        =   1
      Top             =   2130
      Width           =   4890
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando Modulos ...."
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
      Height          =   210
      Left            =   1245
      TabIndex        =   0
      Top             =   3210
      Width           =   2805
   End
   Begin VB.Label lbl_Top 
      BackColor       =   &H008B4914&
      Height          =   1185
      Left            =   30
      TabIndex        =   5
      Top             =   30
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00EAC793&
      Height          =   105
      Left            =   -45
      TabIndex        =   4
      Top             =   1215
      Width           =   5295
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
  Splash.rctStatusBar.Max = 4300
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
 Resume Next
End

End Sub

