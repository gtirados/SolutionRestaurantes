VERSION 5.00
Begin VB.Form FrmClave 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar Password..."
   ClientHeight    =   4995
   ClientLeft      =   2385
   ClientTop       =   1800
   ClientWidth     =   3465
   Icon            =   "FrmClave.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4995
   ScaleWidth      =   3465
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   960
      Picture         =   "FrmClave.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   1300
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Cambiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      Picture         =   "FrmClave.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   1300
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   3375
      Begin VB.TextBox TxtClaveNR 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox TxtClaveN 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TxtClaveA 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Confirme Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Password Nuevo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   855
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Password Anterior:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   135
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label lblbarraos 
      Alignment       =   2  'Center
      BackColor       =   &H008B4914&
      Caption         =   "Solution for Business"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   30
      TabIndex        =   11
      Top             =   4680
      Width           =   3495
   End
   Begin VB.Label LblUsuario 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   6
      Top             =   480
      Width           =   2220
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de Usuario:   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "FrmClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmdAceptar_Click()
If TxtClaveA.Text = "" And TxtClaveN.Text = "" And TxtClaveNR.Text = "" Then
    MsgBox "Datos Incorrectos...!!!", 48, Pub_Titulo
    Azul TxtClaveA, TxtClaveA
    Exit Sub
End If
usu.MoveFirst
Do Until usu.EOF
 If Trim(usu!USU_KEY) = Trim(LK_CODUSU) And Trim(usu!USU_CLAVE) = Trim(TxtClaveA.Text) Then
   GoTo iguales
 End If
 usu.MoveNext
Loop
MsgBox "Clave Arterior .. INCORRECTA... !!!", 48, Pub_Titulo
Azul TxtClaveA, TxtClaveA
Exit Sub
iguales:
If TxtClaveN.Text <> TxtClaveNR.Text Then
  MsgBox "La clave nueva reintentada, no es igual ... !!!", 48, Pub_Titulo
  Azul TxtClaveN, TxtClaveN
  Exit Sub
End If
usu.Edit
usu!USU_CLAVE = TxtClaveN.Text
usu.Update
MsgBox "OK ... Clave Cambiada.", 48, Pub_Titulo
Unload FrmClave

'' graba nueva clave
End Sub

Private Sub Cmdclose_Click()
  Unload FrmClave
End Sub


Private Sub Form_Activate()
LblUsuario.Caption = LK_CODUSU

End Sub

Private Sub TxtClaveA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Azul TxtClaveN, TxtClaveN
End If
End Sub

Private Sub TxtClaveN_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Azul TxtClaveNR, TxtClaveNR
End If
End Sub

Private Sub TxtClaveNR_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 CmdAceptar.SetFocus
End If

End Sub
