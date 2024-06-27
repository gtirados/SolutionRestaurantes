VERSION 5.00
Begin VB.Form frmExaminar 
   Caption         =   "Examinar."
   ClientHeight    =   3330
   ClientLeft      =   4680
   ClientTop       =   1680
   ClientWidth     =   4635
   Icon            =   "frmExaminar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4635
   Begin VB.TextBox txtarchi 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txtdire 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Directorio :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre de Archivo "
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Ruta :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
   End
End
Attribute VB_Name = "frmExaminar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
Dim MiDIRE
If UCase(Trim(txtarchi.text)) = "*.XLS" Then Exit Sub
If Trim(txtdire.text) = "" Then Exit Sub
MiDIRE = txtdire.text
If Trim(Dir(MiDIRE, vbDirectory)) = "" Then
  pub_mensaje = "Directorio No Existe,  Desea Crearlo.. ?"
  Pub_Respuesta = MsgBox(pub_mensaje, vbDefaultButton2 + vbQuestion + vbOKCancel, Pub_Titulo)
  If Pub_Respuesta = vbNo Then
    Exit Sub
  End If
Else
  If UCase(Right(Trim(txtarchi.text), 4)) = ".XLS" Then
     frmDiseno.txtruta.text = txtdire.text + "\" + Trim(txtarchi.text)
  Else
     frmDiseno.txtruta.text = txtdire.text + "\" + Trim(txtarchi.text) + ".XLS"
  End If
  ChDir txtdire.text
  Unload frmExaminar
  Exit Sub
End If
MkDir Trim(txtdire.text)
If UCase(Right(Trim(txtarchi.text), 4)) = ".XLS" Then
     frmDiseno.txtruta.text = txtdire.text + "\" + Trim(txtarchi.text)
Else
     frmDiseno.txtruta.text = txtdire.text + "\" + Trim(txtarchi.text) + ".xls"
End If
ChDir txtdire.text
Unload frmExaminar
End Sub

Private Sub Command1_Click()
Unload frmExaminar
End Sub

Private Sub Dir1_Change()
txtdire.text = Dir1.Path
End Sub

Private Sub Form_Load()
Dim DIREC
txtarchi.text = "*.xls"
txtdire.text = Dir1.Path
End Sub

Private Sub txtarchi_GotFocus()
 Azul txtarchi, txtarchi
End Sub
