VERSION 5.00
Begin VB.Form Splash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2910
   ClientLeft      =   825
   ClientTop       =   1155
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
   ScaleHeight     =   2910
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Label Empresa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Top             =   1770
      Width           =   75
   End
   Begin VB.Label LblMensa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Iniciando..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   150
      TabIndex        =   7
      Top             =   2220
      Width           =   4965
   End
   Begin VB.Label lblporcentaje 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0%..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   2130
      TabIndex        =   5
      Top             =   2580
      Width           =   825
   End
   Begin VB.Label Label1 
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
      Left            =   90
      TabIndex        =   4
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label Label2 
      BackColor       =   &H00EAC793&
      Height          =   105
      Left            =   0
      TabIndex        =   3
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
      Left            =   105
      TabIndex        =   2
      Top             =   90
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
      Height          =   240
      Left            =   1260
      TabIndex        =   1
      Top             =   510
      Width           =   1005
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gestion de Restaurantes y Bares"
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
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   840
      Width           =   3225
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H008B4914&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      Height          =   2835
      Left            =   0
      Top             =   60
      Width           =   5250
   End
   Begin VB.Label lbl_Top 
      BackColor       =   &H008B4914&
      Height          =   1170
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5295
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H008B4914&
      BorderWidth     =   2
      Height          =   1605
      Left            =   15
      Top             =   1290
      Width           =   5250
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
    Pub_ConnAdo.Close
    Screen.MousePointer = 0
    End
End If
Exit Sub
SALE:
End
End Sub
Private Sub Form_Load()
Dim wflag_bloq As String * 1
Dim success%
Dim PB
PB = Chr(10) & Chr(13) & Chr(10) & Chr(13)
'On Error GoTo SALE
Screen.MousePointer = 11
If App.PrevInstance Then
  pub_mensaje = App.Path & " " & "Software"
  pub_mensaje = pub_mensaje & PB & "Posiblemente la Aplicación este cargada o no ha sido cerrada Correctamente "
  pub_mensaje = pub_mensaje & PB & "Debe Cerrar todos los Programas e Iniciar la seccion como Usuario Distinto ..."
  MsgBox pub_mensaje, vbCritical, "Software"
  Screen.MousePointer = 0
  End
End If

Pub_Titulo = "UniSoft S.A.C. - Solution"
LK_CODCIA = ""
LK_CODUSU = ""
If Nulo_Valor0(PUB_FLAG) = 0 Then
  wflag_bloq = ""
  If dir("C:\WINDOWS\Sisgts", vbDirectory) <> "" Then
    wflag_bloq = "A"
  End If
  If dir("C:\Winnt\Sisgts", vbDirectory) <> "" Then
    wflag_bloq = "A"
  End If
  If dir("C:\Win98\Sisgts", vbDirectory) <> "" Then
    wflag_bloq = "A"
  End If
  If wflag_bloq <> "A" Then
    MsgBox "Equipo: MicroProcesador No Identificado..." & Chr(13) & "- Esta copia del Ejecutable no procede - No tiene licencia de uso", vbCritical, "Proveedor del Software - Celular: #990905152"
    End
  End If

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


