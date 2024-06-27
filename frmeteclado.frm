VERSION 5.00
Begin VB.Form frmDetalle 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13050
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   13050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtMesa 
      Alignment       =   2  'Center
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
      Left            =   1320
      TabIndex        =   57
      Top             =   5760
      Width           =   1575
   End
   Begin VB.TextBox txtPax 
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
      Left            =   3360
      TabIndex        =   54
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Frame frameTeclado 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   12855
      Begin VB.CommandButton cmdBorraPax 
         BackColor       =   &H00808080&
         Caption         =   "Cl"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   3000
         Width           =   1575
      End
      Begin VB.CommandButton cmdNum 
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
         Height          =   735
         Index           =   9
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   3000
         Width           =   735
      End
      Begin VB.CommandButton cmdNum 
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
         Height          =   735
         Index           =   8
         Left            =   11880
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton cmdNum 
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
         Height          =   735
         Index           =   7
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton cmdNum 
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
         Height          =   735
         Index           =   6
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton cmdNum 
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
         Height          =   735
         Index           =   5
         Left            =   11880
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmdNum 
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
         Height          =   735
         Index           =   4
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmdNum 
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
         Height          =   735
         Index           =   3
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmdNum 
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
         Height          =   735
         Index           =   2
         Left            =   11880
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdNum 
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
         Height          =   735
         Index           =   1
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdNum 
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
         Height          =   735
         Index           =   0
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   480
         Width           =   735
      End
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
         Height          =   735
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdTecla 
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
         Height          =   735
         Index           =   36
         Left            =   7464
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
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
         Height          =   735
         Index           =   35
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
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
         Height          =   735
         Index           =   34
         Left            =   2586
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
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
         Height          =   735
         Index           =   33
         Left            =   1773
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
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
         Height          =   735
         Index           =   32
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
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
         Height          =   735
         Index           =   31
         Left            =   5025
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
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
         Height          =   735
         Index           =   30
         Left            =   4212
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
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
         Height          =   735
         Index           =   29
         Left            =   3399
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
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
         Height          =   735
         Index           =   28
         Left            =   7464
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
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
         Height          =   735
         Index           =   1
         Left            =   6651
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
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
         Height          =   735
         Index           =   0
         Left            =   5838
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   360
         Width           =   735
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
         Height          =   735
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2880
         Width           =   735
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
         Height          =   735
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3720
         Width           =   5655
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   3399
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   4212
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "f"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   14
         Left            =   3399
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "g"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   15
         Left            =   4212
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2040
         Width           =   735
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
         TabIndex        =   25
         Top             =   360
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
         Height          =   2655
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   27
         Left            =   6651
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "n"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   26
         Left            =   5838
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "b"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   25
         Left            =   5025
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "v"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   24
         Left            =   4212
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "c"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   23
         Left            =   3399
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   22
         Left            =   2586
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "z"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   21
         Left            =   1773
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "ñ"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   20
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "l"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   19
         Left            =   7464
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "k"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   18
         Left            =   6651
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "j"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   17
         Left            =   5838
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "h"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   16
         Left            =   5025
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "d"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   13
         Left            =   2586
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   12
         Left            =   1773
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   11
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "p"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   10
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "o"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   9
         Left            =   7464
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "i"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   8
         Left            =   6651
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "u"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   5838
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "y"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   6
         Left            =   5025
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "e"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   2586
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton cmdTecla 
         Caption         =   "w"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   1773
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton q 
         Caption         =   "q"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PAX"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   11160
         TabIndex        =   56
         Top             =   120
         Width           =   435
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Height          =   3615
         Left            =   10080
         Top             =   360
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const VK_H = 72
Const VK_E = 69
Const VK_L = 76
Const VK_O = 79
Const VK_CR = 13
Const VK_ESC = 27
Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2
Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Sub cmdBorra_Click()
txtMesa.Text = ""
End Sub

Private Sub cmdBorraPax_Click()
txtPax.Text = ""
End Sub

Private Sub cmdEnter_Click()
    txtMesa.SetFocus
    keybd_event VK_CR, 0, 0, 0
End Sub

Private Sub cmdEsc_Click()
    txtMesa.SetFocus
    keybd_event VK_ESC, 0, 0, 0
End Sub

Private Sub cmdEspacio_Click()
txtMesa.Text = txtMesa.Text + Space(1)
End Sub

Private Sub cmdNum_Click(Index As Integer)
txtPax.Text = txtPax.Text + cmdNum(Index).Caption
End Sub

Private Sub cmdRet_Click()
If Len(txtMesa.Text) > 0 Then
    txtMesa.Text = Left(txtMesa.Text, Len(txtMesa.Text) - 1)
End If
End Sub

Private Sub cmdTecla_Click(Index As Integer)
txtMesa.Text = txtMesa.Text + cmdTecla(Index).Caption
End Sub

Private Sub txtmesa_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 13
    MsgBox "Pulsó ENTER"
Case 27
    MsgBox "Pulsó ESC"
End Select
End Sub
