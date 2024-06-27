VERSION 5.00
Begin VB.Form FrmLogo 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Acerca de. . ."
   ClientHeight    =   3210
   ClientLeft      =   1665
   ClientTop       =   2505
   ClientWidth     =   6165
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3210
   ScaleWidth      =   6165
   Begin VB.CommandButton Command2 
      Caption         =   "Retornar"
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Win Soft  : Software desarrollado para control Administrativo Contable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4770
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Telf.: 22 0742 - 29 2859    Trujillo - Perú "
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
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Solutin Peru S.A.C."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   0
      Picture         =   "FrmLogo.frx":0000
      Top             =   0
      Width           =   2550
   End
End
Attribute VB_Name = "FrmLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim va As Integer

Private Sub Command1_Click()
Screen.MousePointer = 11
'OLE1.DoVerb 1
Screen.MousePointer = 0

End Sub


Private Sub Command2_Click()
Unload FrmLogo
End Sub


Private Sub Form_Load()
va = 0
End Sub
