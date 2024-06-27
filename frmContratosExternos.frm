VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmContratosExternos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Datos Opcionales"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9840
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
   ScaleHeight     =   4515
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvData 
      Height          =   3015
      Left            =   1440
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   750
      Visible         =   0   'False
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9720
      Top             =   2160
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
            Picture         =   "frmContratosExternos.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContratosExternos.frx":039A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   635
      ButtonWidth     =   2011
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
            Caption         =   "&Aceptar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtDescExterno 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   7215
   End
   Begin VB.TextBox txtPrecio 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdAgregarE 
      Caption         =   "A&gregar"
      Height          =   600
      Left            =   8760
      Picture         =   "frmContratosExternos.frx":0734
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   990
   End
   Begin VB.CommandButton cmdQuitarE 
      Caption         =   "&Quitar"
      Height          =   600
      Left            =   8760
      Picture         =   "frmContratosExternos.frx":0ABE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   990
   End
   Begin MSComctlLib.ListView lvExternos 
      Height          =   3135
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1200
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   7560
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCIÓN:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD:"
      Height          =   195
      Left            =   6555
      TabIndex        =   8
      Top             =   885
      Width           =   1020
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO:"
      Height          =   195
      Left            =   600
      TabIndex        =   7
      Top             =   885
      Width           =   750
   End
   Begin VB.Menu mnupopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnuEditar 
         Caption         =   "Editar"
      End
   End
End
Attribute VB_Name = "frmContratosExternos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vBuscar   As Boolean 'variable para la busqueda de Productos
Public oRsExt As ADODB.Recordset
Private vPunto As Boolean
Private loc_key   As Integer

Private Sub cmdAgregarE_Click()
AgregarExterno
End Sub

Private Sub cmdQuitarE_Click()
If Me.lvExternos.ListItems.count = 0 Then Exit Sub
Me.lvExternos.ListItems.Remove Me.lvExternos.SelectedItem.Index
Me.cmdQuitarE.Enabled = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
vPunto = False
    Dim Item As Object
ConfiguraLV
If oRsExt.RecordCount > 0 Then
oRsExt.Filter = ""
    oRsExt.MoveFirst
End If
    Do While Not oRsExt.EOF
        Set Item = Me.lvExternos.ListItems.Add(, , Trim(oRsExt!DESCRIPCION))
        Item.Tag = oRsExt!CODEXTERNO
        Item.SubItems(1) = oRsExt!Cantidad
        Item.SubItems(2) = oRsExt!PRECIO
        Item.SubItems(3) = oRsExt!Importe
        oRsExt.MoveNext
    Loop

End Sub

Private Sub lvExternos_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.cmdQuitarE.Enabled = True
End Sub

Private Sub lvExternos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Me.lvExternos.ListItems.count = 0 Then Exit Sub
'If Me.lvExternos.SelectedItem.ListSubItems.count = 0 Then Exit Sub
If Button = 2 Then
PopupMenu mnupopup
End If
End Sub

Private Sub mnuEditar_Click()

    With Me.lvExternos.SelectedItem
        frmContratosExternosEdit.txtDescExterno.Text = .Text
        frmContratosExternosEdit.txtCantidad.Text = .SubItems(1)
        frmContratosExternosEdit.txtPrecio.Text = .SubItems(2)
    End With

    frmContratosExternosEdit.Show vbModal
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index

        Case 1

            Dim Fila As Object

            'If Not frmContratos.oRsExternos.BOF And Not frmContratos.oRsExternos.EOF Then
            If frmContratos.oRsExternos.RecordCount <> 0 Then
            frmContratos.oRsExternos.MoveFirst
            
            Do While Not frmContratos.oRsExternos.EOF
                frmContratos.oRsExternos.Delete
                frmContratos.oRsExternos.MoveNext
            Loop
End If
            frmContratos.oRsExternos.Filter = ""
            If frmContratos.oRsExternos.RecordCount <> 0 Then
                frmContratos.oRsExternos.Filter = ""
                frmContratos.oRsExternos.MoveFirst
            End If
    
            'End If
Dim Vnro As Integer

            For Each Fila In Me.lvExternos.ListItems
'Vnro = frmContratos.oRsAlternativas.RecordCount + 1
                frmContratos.oRsExternos.AddNew
                frmContratos.oRsExternos.Fields!CODEXTERNO = Fila.Tag
                frmContratos.oRsExternos.Fields!DESCRIPCION = Fila.Text
                frmContratos.oRsExternos.Fields!Cantidad = Fila.SubItems(1)
                frmContratos.oRsExternos.Fields!PRECIO = Fila.SubItems(2)
                frmContratos.oRsExternos.Fields!Importe = val(Fila.SubItems(1)) * val(Fila.SubItems(2))
                
                frmContratos.oRsExternos.Update
                'Vnro = Vnro + 1
            Next
            
          
            frmContratos.lblExternos.Caption = frmContratos.oRsExternos.RecordCount & " ITEMS AGREGADOS"
            frmContratos.oRsExternos.Filter = ""
            frmContratos.oRsExternos.MoveFirst
            Unload Me

        Case 2
            Unload Me

    End Select

End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
  If SoloNumeros(KeyAscii) Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then AgregarExterno
End Sub

Private Sub txtDescExterno_Change()
  vBuscar = True
End Sub

Private Sub txtDescExterno_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error GoTo SALE

    Dim strFindMe As String

    Dim itmFound  As Object ' ListItem    ' Variable FoundItem.

    If KeyCode = 40 Then  ' flecha abajo
        loc_key = loc_key + 1

        If loc_key > Me.lvData.ListItems.count Then loc_key = lvData.ListItems.count
        GoTo POSICION
    End If

    If KeyCode = 38 Then
        loc_key = loc_key - 1

        If loc_key < 1 Then loc_key = 1
        GoTo POSICION
    End If

    If KeyCode = 34 Then
        loc_key = loc_key + 17

        If loc_key > lvData.ListItems.count Then loc_key = lvData.ListItems.count
        GoTo POSICION
    End If

    If KeyCode = 33 Then
        loc_key = loc_key - 17

        If loc_key < 1 Then loc_key = 1
        GoTo POSICION
    End If

    If KeyCode = 27 Then
        Me.lvData.Visible = False
        Me.txtDescExterno.Text = ""
        '        Me.lblDocumento.Caption = ""
        '        Me.lblTelefonos.Caption = ""
    End If

    GoTo fin
POSICION:
    lvData.ListItems.Item(loc_key).Selected = True
    lvData.ListItems.Item(loc_key).EnsureVisible
    'txtRS.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
    Me.txtDescExterno.SelStart = Len(Me.txtDescExterno.Text)
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtDescExterno_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then
        If vBuscar Then
            Me.lvData.ListItems.Clear
            LimpiaParametros oCmdEjec

            Dim orsDATA As ADODB.Recordset

            oCmdEjec.CommandText = "SPFILTRAPLATOS"
            Set orsDATA = oCmdEjec.Execute(, Array(LK_CODCIA, Me.txtDescExterno.Text))

            Dim Item As Object
        
            If Not orsDATA.EOF Then

                Do While Not orsDATA.EOF
                    Set Item = Me.lvData.ListItems.Add(, , Trim(orsDATA!PRODUCTO))
                    Item.Tag = Trim(orsDATA!Codigo)
                    Item.SubItems(1) = orsDATA!PRECIO
                    orsDATA.MoveNext
                Loop

                Me.lvData.Visible = True
                Me.lvData.ListItems(1).Selected = True
                loc_key = 1
                Me.lvData.ListItems(1).EnsureVisible
                vBuscar = False
            
                '         If MsgBox("Cliente no existe." + vbCrLf + "¿Desea Crearlo.?", vbQuestion + vbYesNo, "Restaurantes") = vbYes Then
                '         frmCLI.Show vbModal
                '         End If
                Else
                Me.lvData.Visible = False
            End If

        Else

            If Me.txtDescExterno.Text <> "" Then
                Me.lvData.Visible = False
                Me.txtDescExterno.Text = Me.lvData.SelectedItem.Text
                Me.txtDescExterno.Tag = Me.lvData.SelectedItem.Tag
                Me.txtPrecio.Text = Me.lvData.SelectedItem.SubItems(1)
                'Me.FraDetalle.Enabled = True
                Me.txtCantidad.SetFocus
                Me.txtCantidad.SelStart = 0
                Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
            End If
        End If
    End If

End Sub

Private Sub txtPrecio_Change()
 If InStr(txtPrecio.Text, ".") = 0 Then
        vPunto = False
    Else
        vPunto = True
    End If
End Sub

Private Sub txtPrecio_GotFocus()
 If InStr(txtPrecio.Text, ".") = 0 Then
        vPunto = False
    Else
        vPunto = True
    End If
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
If NumerosyPunto(KeyAscii) Then KeyAscii = 0
    If Chr(KeyAscii) = "." Then
        If vPunto = True Or Len(Trim(txtPrecio.Text)) = 0 Then
         
            KeyAscii = 0
         
        End If
    End If
If KeyAscii = vbKeyReturn Then Me.txtCantidad.SetFocus
End Sub

Private Sub AgregarExterno()

    Dim vPasa As Boolean

    vPasa = True

    If Len(Trim(Me.txtDescExterno.Text)) = 0 Then
        MsgBox "Debe ingresar la Descripción.", vbCritical, Pub_Titulo
        Me.txtDescExterno.SetFocus
    ElseIf Len(Trim(Me.txtCantidad.Text)) = 0 Then
        MsgBox "Debe ingresar la Cantidad.", vbCritical, Pub_Titulo
        Me.txtCantidad.SetFocus
    ElseIf val(Me.txtCantidad.Text) = 0 Or Not IsNumeric(Me.txtCantidad.Text) Then
        MsgBox "La Cantidad proporcionada es incorrecta.", vbCritical, Pub_Titulo
        Me.txtCantidad.SetFocus
        Me.txtCantidad.SelStart = 0
        Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
    ElseIf Len(Trim(Me.txtPrecio.Text)) = 0 Then
        MsgBox "Debe ingresar el Precio.", vbCritical, Pub_Titulo
        Me.txtPrecio.SetFocus
    ElseIf val(Me.txtPrecio.Text) = 0 Or Not IsNumeric(Me.txtPrecio.Text) Then
        MsgBox "El precio proporcinado es incorrecto.", vbCritical, Pub_Titulo
        Me.txtPrecio.SetFocus
        Me.txtPrecio.SelStart = 0
        Me.txtPrecio.SelLength = Len(Me.txtPrecio.Text)
    Else

        Dim Fila As Object

        If Not frmContratos.oRsAlternativas Is Nothing Then
            frmContratos.oRsAlternativas.Filter = ""

        End If

        If Me.lvExternos.ListItems.count = 0 Then
            Set Fila = Me.lvExternos.ListItems.Add(, , Trim(Me.txtDescExterno.Text))
            Fila.Tag = Me.txtDescExterno.Tag
            Fila.SubItems(1) = Me.txtCantidad.Text
            Fila.SubItems(2) = Me.txtPrecio.Text
            Fila.SubItems(3) = val(Me.txtCantidad.Text) * val(Me.txtPrecio.Text)
            Me.txtDescExterno.Text = ""
            Me.txtCantidad.Text = ""
            Me.txtPrecio.Text = ""
            Me.txtDescExterno.SetFocus
        Else

            For Each Fila In Me.lvExternos.ListItems

                If Fila.Tag = Me.txtDescExterno.Tag Then
                    vPasa = False
                    Exit For
                End If
            Next
    
            If vPasa Then
                Set Fila = Me.lvExternos.ListItems.Add(, , Trim(Me.txtDescExterno.Text))
                Fila.Tag = Me.txtDescExterno.Tag
                Fila.SubItems(1) = Me.txtCantidad.Text
                Fila.SubItems(2) = Me.txtPrecio.Text
                Fila.SubItems(3) = val(Me.txtCantidad.Text) * val(Me.txtPrecio.Text)
                Me.txtDescExterno.Text = ""
                Me.txtCantidad.Text = ""
                Me.txtPrecio.Text = ""
                Me.txtDescExterno.SetFocus
            Else
                MsgBox "El Producto proporcionado ya se encuentra en la lista.", vbCritical, Pub_Titulo
                Me.txtDescExterno.SetFocus
                Me.txtDescExterno.SelStart = 0
                Me.txtDescExterno.SelLength = Len(Me.txtDescExterno.Text)
            End If
    
        End If
        
    End If

End Sub

Private Sub ConfiguraLV()

    With Me.lvData
        .ColumnHeaders.Add , , "DESCRIPCION", 4000
        .ColumnHeaders.Add , , "PRECIO", 1000
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
    End With

    With Me.lvExternos
        .ColumnHeaders.Add , , "DESCRIPCION", 4000
        .ColumnHeaders.Add , , "CANTIDAD", 1000
        .ColumnHeaders.Add , , "PRECIO", 1000
        .ColumnHeaders.Add , , "IMPORTE", 1000
        .HideColumnHeaders = False
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
    End With

End Sub
