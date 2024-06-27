VERSION 5.00
Begin VB.Form frmFaccomandaOtroPlato 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingrese Producto"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6105
   BeginProperty Font 
      Name            =   "Tahoma"
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
   ScaleHeight     =   1530
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   600
      Left            =   4680
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   600
      Left            =   3240
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtProducto 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "frmFaccomandaOtroPlato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
If Len(Trim(Me.txtProducto.Text)) <> 0 Then
frmFacComanda.lvDetalle.SelectedItem.SubItems(2) = Me.txtProducto.Text
frmFacComanda.lvDetalle.SelectedItem.SubItems(13) = Me.txtProducto.Text
Unload Me
End If
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub
