VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmComandaProductoSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscador de Productos"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12240
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
   ScaleHeight     =   7230
   ScaleWidth      =   12240
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvData 
      Height          =   6495
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   11456
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
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   10815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCTO:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1065
   End
End
Attribute VB_Name = "frmComandaProductoSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private loc_key As Integer
Private posicion As Integer
Public gDELIVERY As Boolean
Public gMostrador As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
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
    If frmComandaProductoSearch.gDELIVERY = True Then
    oCmdEjec.CommandText = "SPPRODUCTOS_SEARCH_DELIVERY"
    Else
    oCmdEjec.CommandText = "SPPRODUCTOS_SEARCH"
    End If
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("SEARCH", adVarChar, adParamInput, 80, Me.txtSearch.Text)

    Dim ORSd As ADODB.Recordset

    Set ORSd = oCmdEjec.Execute

    If Not ORSd.EOF Then
        loc_key = 1

        Do While Not ORSd.EOF
            Set itemX = Me.lvData.ListItems.Add(, , ORSd!Codigo)
            itemX.SubItems(1) = ORSd!plato
            itemX.SubItems(2) = ORSd!PRECIO
            itemX.Tag = ORSd!Familia
            ORSd.MoveNext
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
        GoTo posicion
    End If

    If KeyCode = 38 Then
        loc_key = loc_key - 1

        If loc_key < 1 Then loc_key = 1
        GoTo posicion
    End If

    If KeyCode = 34 Then
        loc_key = loc_key + 17

        If loc_key > lvData.ListItems.count Then loc_key = lvData.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 33 Then
        loc_key = loc_key - 17

        If loc_key < 1 Then loc_key = 1
        GoTo posicion
    End If

    If KeyCode = 27 Then
        Me.lvData.Visible = False
'        Me.txtCliente.Text = ""
'        Me.lblDocumento.Caption = ""
'        Me.lblTelefonos.Caption = ""
    End If

    GoTo fin
posicion:
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
            EnviarSeleccion
            Me.txtSearch.SelStart = 0
            Me.txtSearch.SelLength = Len(Me.txtSearch.Text)
        End If
    End If

End Sub

Private Sub EnviarSeleccion()

    If gDELIVERY Then
        frmDeliveryApp.vCodFam = Me.lvData.SelectedItem.Tag
        frmDeliveryApp.AgregarDesdeBuscador Me.lvData.SelectedItem.Text, Me.lvData.SelectedItem.SubItems(1), Me.lvData.SelectedItem.SubItems(2)
    Else

        If gMostrador Then
            frmComanda2.vCodFam = Me.lvData.SelectedItem.Tag
            frmComanda2.AgregarDesdeBuscador Me.lvData.SelectedItem.Text, Me.lvData.SelectedItem.SubItems(1), Me.lvData.SelectedItem.SubItems(2)
        Else
            frmComanda.vCodFam = Me.lvData.SelectedItem.Tag
            frmComanda.AgregarDesdeBuscador Me.lvData.SelectedItem.Text, Me.lvData.SelectedItem.SubItems(1), Me.lvData.SelectedItem.SubItems(2)
        End If
    End If


End Sub

