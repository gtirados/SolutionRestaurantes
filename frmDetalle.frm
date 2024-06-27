VERSION 5.00
Begin VB.Form frmDetalle 
   BackColor       =   &H8000000C&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalle Adicional de Plato"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9240
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcliente 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   45
      Top             =   120
      Visible         =   0   'False
      Width           =   9015
   End
   Begin VB.TextBox txtMesa 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   9015
   End
   Begin VB.Frame frameTeclado 
      BackColor       =   &H8000000C&
      Height          =   3735
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   9015
      Begin VB.CommandButton cmdRet 
         BackColor       =   &H00808080&
         Caption         =   "Ret"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   ","
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   37
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         BackColor       =   &H00C0C0C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         BackColor       =   &H00C0C0C0&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   5
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         BackColor       =   &H00C0C0C0&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         BackColor       =   &H00C0C0C0&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         BackColor       =   &H00C0C0C0&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   8
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         BackColor       =   &H00C0C0C0&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         BackColor       =   &H00C0C0C0&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdBorra 
         BackColor       =   &H00808080&
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdEspacio 
         BackColor       =   &H00808080&
         Caption         =   "Espacio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3120
         Width           =   5655
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   13
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   14
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   23
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   24
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdEsc 
         BackColor       =   &H00808080&
         Caption         =   "Esc"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdEnter 
         BackColor       =   &H00808080&
         Caption         =   "Enter"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   36
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   35
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   34
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   33
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   32
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   31
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   30
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "Ñ"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   29
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   28
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "K"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   27
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   26
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   25
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   22
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   21
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   20
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   19
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   18
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   17
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   16
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   15
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   12
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   11
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdtecla 
         Caption         =   "Q"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   10
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   960
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vDetalle As String
Public vSelec As Boolean
Public EsDetalle As Boolean
Public ParaCliente As Boolean
Public Comensales As Boolean
Option Explicit
'
'Const VK_H = 72
'Const VK_E = 69
'Const VK_L = 76
'Const VK_O = 79
'Const VK_CR = 13
'Const VK_ESC = 27
'Const KEYEVENTF_EXTENDEDKEY = &H1
'Const KEYEVENTF_KEYUP = &H2
'Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Sub cmdBorra_Click()

    If EsDetalle = True Then
       Me.txtCliente.Text = ""
    Else
         txtMesa.Text = ""
    End If

End Sub

Private Sub AsignaValor(vCli As Boolean)

    vSelec = True

    If EsDetalle = True Then
         vDetalle = Trim(Me.txtCliente.Text)
    Else
       vDetalle = Trim(Me.txtMesa.Text)
    End If

    Unload Me
End Sub

Private Sub cmdEnter_Click()

    AsignaValor ParaCliente

End Sub

Private Sub cmdEsc_Click()

    If EsDetalle = True Then
     Me.txtCliente.Text = ""
        Me.txtCliente.SetFocus
    Else
      txtMesa.Text = ""
        txtMesa.SetFocus
    End If

End Sub

Private Sub cmdEspacio_Click()

    If EsDetalle = True Then
      txtCliente.Text = txtCliente.Text + Space(1)
    Else
          txtMesa.Text = txtMesa.Text + Space(1)
    End If

End Sub

Private Sub cmdRet_Click()

    If EsDetalle = True Then
        If Len(Me.txtCliente.Text) > 0 Then
            Me.txtCliente.Text = Left(Me.txtCliente.Text, Len(Me.txtCliente.Text) - 1)
        End If
      
    Else
  If Len(txtMesa.Text) > 0 Then
            txtMesa.Text = Left(txtMesa.Text, Len(txtMesa.Text) - 1)
        End If

    End If

End Sub

Private Sub cmdTecla_Click(Index As Integer)

    If EsDetalle = True Then
        txtCliente.Text = txtCliente.Text + cmdtecla(Index).Caption
    Else
        txtMesa.Text = txtMesa.Text + cmdtecla(Index).Caption
    End If

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    vSelec = False

    If EsDetalle Then
        Me.txtCliente.MaxLength = 120
        Me.txtMesa.Visible = False
        Me.txtCliente.Visible = True
    Else
     Me.txtMesa.MaxLength = 120
        Me.txtCliente.Visible = False
        Me.txtMesa.Visible = True
    End If

    Dim i As Integer

    If Comensales Then

        For i = 10 To 37
            Me.cmdtecla(i).Enabled = False
        Next

    Else

        For i = 10 To 37
            Me.cmdtecla(i).Enabled = True
        Next

    End If

End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then AsignaValor ParaCliente
End Sub

Private Sub txtmesa_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then AsignaValor ParaCliente
End Sub
