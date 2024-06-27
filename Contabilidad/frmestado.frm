VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmestado 
   Caption         =   "Form1"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   8505
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   3836
      _Version        =   327680
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      Begin VB.OptionButton opes 
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton opes 
         Caption         =   "Cliente "
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmestado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
