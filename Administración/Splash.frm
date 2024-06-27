VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Splash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F5F1EC&
   BorderStyle     =   0  'None
   ClientHeight    =   2940
   ClientLeft      =   825
   ClientTop       =   1155
   ClientWidth     =   5235
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
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2940
   ScaleWidth      =   5235
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.ProgressBar rctStatusBar 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      BackColor       =   &H00EAC793&
      Height          =   105
      Index           =   2
      Left            =   0
      TabIndex        =   9
      Top             =   1170
      Width           =   5295
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Solution"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1650
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
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   1200
      TabIndex        =   6
      Top             =   420
      Width           =   1050
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin Solution"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   3180
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Empresa 
      BackStyle       =   0  'Transparent
      Caption         =   "Módulo de Administración"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1920
      Width           =   4095
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
      Caption         =   "Cargando Modulos ...."
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
   Begin VB.Label lbl_Top 
      BackColor       =   &H008B4914&
      Height          =   1170
      Left            =   0
      TabIndex        =   8
      Top             =   0
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
Dim PB
PB = Chr(10) & Chr(13) & Chr(10) & Chr(13)
On Error GoTo SALE
CenterMe Splash
Screen.MousePointer = 11
If App.PrevInstance Then
  pub_mensaje = App.Path & " " & "SOLUTIN"
  pub_mensaje = pub_mensaje & PB & "Posiblemente la Aplicación este cargada o no ha sido cerrada Correctamente "
  pub_mensaje = pub_mensaje & PB & "Debe Cerrar todos los Programas e Iniciar la seccion como Usuario Distinto ..."
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
End

End Sub

