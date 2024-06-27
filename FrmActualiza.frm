VERSION 5.00
Begin VB.Form FrmActualiza 
   Caption         =   "Actualizacion  de Transacciones (FORMGEN)"
   ClientHeight    =   2445
   ClientLeft      =   2175
   ClientTop       =   1830
   ClientWidth     =   5385
   Icon            =   "FrmActualiza.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   5385
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   -120
      ScaleHeight     =   1275
      ScaleWidth      =   6075
      TabIndex        =   2
      Top             =   1800
      Width           =   6135
      Begin VB.CommandButton Command2 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   3480
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton CmdIniciar 
         Caption         =   "Iniciar Proceso  "
         Height          =   375
         Left            =   480
         TabIndex        =   0
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"FrmActualiza.frx":0442
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "FrmActualiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdIniciar_Click()
ACT_FORMGEN
'Unload FORMGEN
End Sub

Private Sub Command2_Click()
  Unload FrmActualiza
End Sub

