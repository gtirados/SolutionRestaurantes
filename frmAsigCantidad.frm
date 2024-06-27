VERSION 5.00
Begin VB.Form frmAsigCantidad 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asignar Cantidad"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2895
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
   ScaleHeight     =   1275
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CANTIDAD:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1020
   End
End
Attribute VB_Name = "frmAsigCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vPUNTO As Boolean
Public xtipo As String

Private Sub Enviar()
If Len(Trim(Me.txtCantidad.Text)) = 0 Then
    MsgBox "Debe ingresar la cantidad", vbCritical, "Error de Datos"
    Me.txtCantidad.SetFocus
    Exit Sub
End If
If val(Me.txtCantidad.Text) = 0 Then
    MsgBox "Cantidad proporcionada incorrecta", vbCritical, "Error de Datos"
    Me.txtCantidad.SetFocus
    Exit Sub
End If

Dim itemX As Object
Dim vPasa As Boolean
vPasa = True
If xtipo = "M" Then 'se agrega a la formulación del plato

If frmListaProd.cboTipoProd.ListIndex = 3 Then 'materia prima

Dim vTotPro As Double
vTotPro = frmProductos.txtProporcion.Text

Dim itemI As Object
Dim totA As Double
totA = 0
For Each itemI In frmProductos.lvFormulacion.ListItems
totA = totA + itemI.SubItems(3)
    Next
    totA = totA + Me.txtCantidad.Text
'    If totA > vTotPro Then    quitado gts
'        MsgBox "Las cantidades no coincidem.", vbInformation, Pub_Titulo
'        Exit Sub
'    End If
    
    End If
    
    If frmProductos.lvFormulacion.ListItems.count = 0 Then
        Set itemX = frmProductos.lvFormulacion.ListItems.Add(, , frmBusProd.lvData.SelectedItem.Text)
        itemX.SubItems(1) = frmBusProd.lvData.SelectedItem.SubItems(1)
        itemX.SubItems(2) = frmBusProd.lvData.SelectedItem.SubItems(2)
        itemX.SubItems(3) = Me.txtCantidad.Text
        Unload Me
    Else
        For Each itemX In frmProductos.lvFormulacion.ListItems
            If itemX = frmBusProd.lvData.SelectedItem.Text Then
                vPasa = False
                Exit For
            End If
        Next
        If vPasa Then
            Set itemX = frmProductos.lvFormulacion.ListItems.Add(, , frmBusProd.lvData.SelectedItem.Text)
            itemX.SubItems(1) = frmBusProd.lvData.SelectedItem.SubItems(1)
            itemX.SubItems(2) = frmBusProd.lvData.SelectedItem.SubItems(2)
            itemX.SubItems(3) = Me.txtCantidad.Text
            Unload Me
        Else
            MsgBox "El Insumo ya se encuentra en la Preparación.", vbInformation, "Error"
        End If
    
    
    End If
ElseIf xtipo = "P" Then 'se agrega a la composición del combo
 If frmProductos.lvComposicion.ListItems.count = 0 Then
        Set itemX = frmProductos.lvComposicion.ListItems.Add(, , frmBusProd.lvData.SelectedItem.Text)
        itemX.SubItems(1) = frmBusProd.lvData.SelectedItem.SubItems(1)
        itemX.SubItems(2) = frmBusProd.lvData.SelectedItem.SubItems(2)
        itemX.SubItems(3) = Me.txtCantidad.Text
        Unload Me
    Else
        For Each itemX In frmProductos.lvComposicion.ListItems
            If itemX = frmBusProd.lvData.SelectedItem.Text Then
                vPasa = False
                Exit For
            End If
        Next
        If vPasa Then
            Set itemX = frmProductos.lvComposicion.ListItems.Add(, , frmBusProd.lvData.SelectedItem.Text)
            itemX.SubItems(1) = frmBusProd.lvData.SelectedItem.SubItems(1)
            itemX.SubItems(2) = frmBusProd.lvData.SelectedItem.SubItems(2)
            itemX.SubItems(3) = Me.txtCantidad.Text
            Unload Me
        Else
            MsgBox "El Insumo ya se encuentra en la Preparación.", vbInformation, "Error"
        End If
    End If
End If
End Sub

Private Sub cmdAceptar_Click()
Enviar
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtCantidad_Change()
If InStr(Me.txtCantidad.Text, ".") Then
    vPUNTO = True
Else
    vPUNTO = False
End If
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If NumerosyPunto(KeyAscii) Then KeyAscii = 0
 If KeyAscii = 46 Then
    If vPUNTO Or Len(Trim(Me.txtCantidad.Text)) = 0 Then
        KeyAscii = 0
    End If
    End If
    
If KeyAscii = vbKeyReturn Then cmdAceptar_Click

End Sub
