VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConversionEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sub Materia Prima Seleccionada"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4380
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   635
      ButtonWidth     =   1931
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aceptar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3240
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConversionEdit.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConversionEdit.frx":039A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtCantidad 
      Height          =   315
      Left            =   3000
      TabIndex        =   6
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
      Height          =   195
      Left            =   3000
      TabIndex        =   7
      Top             =   1080
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad:"
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   660
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proporción:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Materia Prima:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1260
   End
   Begin VB.Label lblunidad 
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   1395
   End
   Begin VB.Label lblProporcion 
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1395
   End
   Begin VB.Label lblMP 
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4155
   End
End
Attribute VB_Name = "frmConversionEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vCant As Double

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
Me.lblMP.Caption = frmConversion.lvSMP.SelectedItem.Text
Me.lblProporcion.Caption = frmConversion.lvSMP.SelectedItem.SubItems(1)
Me.lblUnidad.Caption = frmConversion.lvSMP.SelectedItem.SubItems(2)
Me.txtCantidad.Text = frmConversion.lvSMP.SelectedItem.SubItems(3)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index

        Case 1

           EnviaCantidad
          
        Case 2
            Unload Me

    End Select

End Sub

Private Sub txtCantidad_GotFocus()
Me.txtCantidad.SelStart = 0
Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If SoloNumeros(KeyAscii) Then KeyAscii = 0
If KeyAscii = vbKeyReturn Then
    EnviaCantidad
    
End If
End Sub

Private Sub EnviaCantidad()
 If Len(Trim(Me.txtCantidad.Text)) = 0 Then
                MsgBox "Debe ingresar la Cantidad.", vbInformation, Pub_Titulo
                Me.txtCantidad.SetFocus

                Exit Sub

            End If

            If val(Me.txtCantidad.Text) = 0 Then
                MsgBox "La Cantidad ingresada es incorrecta.", vbInformation, Pub_Titulo
                Me.txtCantidad.SetFocus
                Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)

                Exit Sub

            End If

            If Not IsNumeric(Me.txtCantidad.Text) Then
                MsgBox "La Cantidad Ingresada es incorrecta." & vbCrLf & "Debe ser un número.", vbInformation, Pub_Titulo
                Me.txtCantidad.SetFocus
                Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)

                Exit Sub

            End If

            vCant = Me.txtCantidad.Text
              Unload Me

End Sub
