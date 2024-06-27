VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPedidoMediaPorcion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medias Porciones"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11085
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
   ScaleHeight     =   6750
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvData 
      Height          =   5655
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   9975
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   9960
      TabIndex        =   3
      Top             =   6240
      Width           =   990
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   8760
      TabIndex        =   2
      Top             =   6240
      Width           =   990
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   9855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Producto"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   750
   End
End
Attribute VB_Name = "frmPedidoMediaPorcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private loc_key As Integer
Private POSICION As Integer
Public gIDproductoOriginal As Integer
Public gIDproducto As Integer
Public gPRODUCTO As String
Public gPRECIO As Double
Public gACEPTAR As Boolean

Private Sub cmdAceptar_Click()
EnviarSeleccion
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
gACEPTAR = False
With Me.lvData
    .Gridlines = True
    .LabelEdit = lvwManual
    .FullRowSelect = True
    .View = lvwReport
    
    .HideSelection = False
    .ColumnHeaders.Add , , "Codigo", 1000
    .ColumnHeaders.Add , , "Plato", 4500
    .ColumnHeaders.Add , , "Precio", 1500
    
    .MultiSelect = False
End With

End Sub

Private Sub lvData_DblClick()
EnviarSeleccion
End Sub

Private Sub txtSearch_Change()
    Me.lvData.ListItems.Clear
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_MEDIAS_PORCIONES"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("SEARCH", adVarChar, adParamInput, 80, Me.txtSearch.Text)

    Dim orsD As ADODB.Recordset

    Set orsD = oCmdEjec.Execute

    If Not orsD.EOF Then
        loc_key = 1
        
        Do While Not orsD.EOF
            Set itemX = Me.lvData.ListItems.Add(, , orsD!IDE)
            itemX.SubItems(1) = orsD!PRODUCTO
            itemX.SubItems(2) = orsD!PRECIO
            orsD.MoveNext
        Loop

    Else
        loc_key = 0
    End If

End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo SALE

    Dim strFindMe As String

    Dim itmFound  As Object ' ListItem    ' Variable FoundItem.

    If KeyCode = 40 Then  ' flecha abajo
        loc_key = loc_key + 1

        If loc_key > lvData.ListItems.count Then loc_key = lvData.ListItems.count
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

''    If KeyCode = 27 Then
''        Me.lvData.Visible = False
''        '        Me.txtCliente.Text = ""
''        '        Me.lblDocumento.Caption = ""
''        '        Me.lblTelefonos.Caption = ""
''    End If

    GoTo fin
POSICION:
    lvData.ListItems.Item(loc_key).Selected = True
    lvData.ListItems.Item(loc_key).EnsureVisible
    'txtRS.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
    Me.txtSearch.SelStart = Len(Me.txtSearch.Text)
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then
        If loc_key <> 0 Then
           
            Me.txtSearch.SelStart = 0
            Me.txtSearch.SelLength = Len(Me.txtSearch.Text)
             EnviarSeleccion
        End If
    End If

End Sub

Private Sub EnviarSeleccion()

    gIDproducto = Me.lvData.SelectedItem.Text
    gPRECIO = Me.lvData.SelectedItem.SubItems(2)
    gPRODUCTO = Me.lvData.SelectedItem.SubItems(1)
    
    If gIDproducto = gIDproductoOriginal Then
        MsgBox "El Producto elegido es el mismo que el seleccionado en el Pedido.", vbInformation, Pub_Titulo
        Exit Sub
    End If

    gACEPTAR = True
    Unload Me
End Sub

