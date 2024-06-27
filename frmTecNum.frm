VERSION 5.00
Begin VB.Form frmTecNum 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Teclado Numerico"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2370
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   2370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdC 
      Caption         =   "C"
      Height          =   495
      Left            =   1560
      TabIndex        =   11
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdPunto 
      Caption         =   "."
      Height          =   495
      Left            =   840
      TabIndex        =   10
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "0"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "3"
      Height          =   495
      Index           =   3
      Left            =   1560
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "2"
      Height          =   495
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "1"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "6"
      Height          =   495
      Index           =   6
      Left            =   1560
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "5"
      Height          =   495
      Index           =   5
      Left            =   840
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "4"
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "9"
      Height          =   495
      Index           =   9
      Left            =   1560
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "8"
      Height          =   495
      Index           =   8
      Left            =   840
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "7"
      Height          =   495
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmTecNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
