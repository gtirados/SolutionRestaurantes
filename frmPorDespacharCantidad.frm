VERSION 5.00
Begin VB.Form frmPorDespacharCantidad 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Despachar Items"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4035
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
   ScaleHeight     =   5670
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "<="
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   2400
      TabIndex        =   14
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdNumero 
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
      Height          =   840
      Index           =   9
      Left            =   2400
      TabIndex        =   13
      Tag             =   "9"
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdNumero 
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
      Height          =   840
      Index           =   8
      Left            =   1560
      TabIndex        =   12
      Tag             =   "8"
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdNumero 
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
      Height          =   840
      Index           =   7
      Left            =   720
      TabIndex        =   11
      Tag             =   "7"
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdNumero 
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
      Height          =   840
      Index           =   6
      Left            =   2400
      TabIndex        =   10
      Tag             =   "6"
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdNumero 
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
      Height          =   840
      Index           =   5
      Left            =   1560
      TabIndex        =   9
      Tag             =   "5"
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdNumero 
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
      Height          =   840
      Index           =   4
      Left            =   720
      TabIndex        =   8
      Tag             =   "4"
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdNumero 
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
      Height          =   840
      Index           =   3
      Left            =   2400
      TabIndex        =   7
      Tag             =   "3"
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdNumero 
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
      Height          =   840
      Index           =   2
      Left            =   1560
      TabIndex        =   6
      Tag             =   "2"
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdNumero 
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
      Height          =   840
      Index           =   1
      Left            =   720
      TabIndex        =   5
      Tag             =   "0"
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdNumero 
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
      Height          =   840
      Index           =   0
      Left            =   720
      TabIndex        =   4
      Tag             =   "0"
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2040
      TabIndex        =   3
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   720
      TabIndex        =   2
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblProducto 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   3795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese Cantidad:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   2220
   End
End
Attribute VB_Name = "frmPorDespacharCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vDespachado As Boolean
Public vCANTIDAD As Double

Private Sub cmdAceptar_Click()

    If Not IsNumeric(frmPorDespacharCantidad.txtCantidad.Text) Then
        MsgBox "Cantidad incorrecta.", vbCritical, Pub_Titulo

        Exit Sub

    End If

    If val(Me.txtCantidad.Text) > val(vCANTIDAD) Then
        MsgBox "No se puede despachar mas de la cantidad de la comanda.", vbCritical, Pub_Titulo

        Exit Sub

    End If
    
    If val(Me.txtCantidad.Text) <= 0 Then
        MsgBox "La Cantidad es incorrecta.", vbCritical, Pub_Titulo

        Exit Sub

    End If

    vCANTIDAD = Me.txtCantidad.Text
    vDespachado = True
    Unload Me
End Sub

Private Sub cmdBorrar_Click()
If Len(Trim(Me.txtCantidad.Text)) = 0 Then Exit Sub
Me.txtCantidad.Text = Left(Me.txtCantidad.Text, Len(Me.txtCantidad.Text) - 1)
End Sub

Private Sub cmdCancelar_Click()
vDespachado = False
Unload Me
End Sub

Private Sub cmdNumero_Click(Index As Integer)

    If Me.txtCantidad.SelLength <> 0 Then

        Dim suno As String, sdos As String

        suno = Mid(Me.txtCantidad.Text, 1, Me.txtCantidad.SelStart)
        sdos = Mid(Me.txtCantidad.Text, Me.txtCantidad.SelStart + Me.txtCantidad.SelLength + 1, Len(Me.txtCantidad.Text))
        Me.txtCantidad.Text = CStr(suno) + CStr(Index) + CStr(sdos)

    Else
        Me.txtCantidad.Text = Me.txtCantidad.Text & Index
    End If

End Sub

Private Sub Form_Activate()
Me.txtCantidad.SelStart = 0
Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
End Sub

Private Sub Form_Load()
Me.txtCantidad.Text = vCANTIDAD
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdAceptar_Click
End Sub
