VERSION 5.00
Begin VB.Form FrmMsgbox 
   BackColor       =   &H008B4914&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Servicio de mensajeria"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7170
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdenvio 
      Caption         =   "Enviar"
      Height          =   735
      Left            =   6360
      Picture         =   "FrmMsgbox.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtmsg 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Width           =   5055
   End
   Begin VB.ListBox user 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1080
      Left            =   1080
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackColor       =   &H008B4914&
      Caption         =   "Mensaje :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H008B4914&
      Caption         =   "Avisar a :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "FrmMsgbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdenvio_Click()

For fila = 0 To user.ListCount - 1
 user.ListIndex = fila
 If user.Selected(fila) Then
  PSUSU_LLAVE(0) = Left(user.Text, 10)
  usu_llave.Requery
  usu_llave.Edit
  usu_llave!USU_MSGBOX = LK_CODUSU & " >> " & txtmsg.Text
  usu_llave.Update
 End If
Next
MsgBox "Mensaje Enviado.", 48, "Mensajeria OSBusiness"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 Unload FrmMsgbox
 Exit Sub
End If

End Sub

Private Sub Form_Load()
CenterMe FrmMsgbox
Dim WCODIGO As String * 10
usu.MoveFirst
Do Until usu.EOF
 WCODIGO = Trim(usu!USU_KEY)
 user.AddItem WCODIGO & " : " & Trim(usu!USU_NOMBRE)
usu.MoveNext
Loop

End Sub

Private Sub txtmsg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  cmdenvio_Click
End If
End Sub

Private Sub user_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  txtmsg.SetFocus
End If
End Sub
