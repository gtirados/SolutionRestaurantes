VERSION 5.00
Begin VB.Form FrmClave 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FAEFDA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar Clave"
   ClientHeight    =   3090
   ClientLeft      =   2385
   ClientTop       =   1800
   ClientWidth     =   3885
   Icon            =   "FrmClave.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3090
   ScaleWidth      =   3885
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2340
      TabIndex        =   4
      Top             =   2565
      Width           =   1125
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Cambiar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   420
      TabIndex        =   3
      Top             =   2550
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Cambiar :"
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
      Height          =   1815
      Left            =   75
      TabIndex        =   7
      Top             =   585
      Width           =   3735
      Begin VB.TextBox TxtClaveNR 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox TxtClaveN 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TxtClaveA 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Repita Clave :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clave Nueva :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clave Anterio :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1230
      End
   End
   Begin VB.Label LblUsuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1680
      TabIndex        =   6
      Top             =   240
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario :  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   210
      TabIndex        =   5
      Top             =   240
      Width           =   975
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
 If Trim(usu!usu_key) = Trim(LK_CODUSU) And Trim(usu!USU_CLAVE) = Trim(TxtClaveA.Text) Then
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

Private Sub cmdClose_Click()
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
