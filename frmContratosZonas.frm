VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmContratosZonas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asignar Zonas a Contrato"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7095
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
   ScaleHeight     =   3435
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "&Quitar"
      Height          =   360
      Left            =   6000
      TabIndex        =   4
      Top             =   2400
      Width           =   990
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Height          =   360
      Left            =   6000
      TabIndex        =   3
      Top             =   1920
      Width           =   990
   End
   Begin VB.TextBox txtNINIOS 
      Height          =   285
      Left            =   4680
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtADULTOS 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo DatZonas 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9840
      Top             =   1440
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
            Picture         =   "frmContratosZonas.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContratosZonas.frx":039A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
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
   Begin MSComctlLib.ListView lvZonas 
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1200
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NIÑOS:"
      Height          =   195
      Left            =   3960
      TabIndex        =   8
      Top             =   885
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ADULTOS:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   885
      Width           =   900
   End
End
Attribute VB_Name = "frmContratosZonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vAcepta As Boolean
Public orsMarcados As ADODB.Recordset

Private Sub ConfiguraLV()
With Me.lvZonas
    .ColumnHeaders.Add , , "ZONA", 2800
    .ColumnHeaders.Add , , "AULTOS", 1000
    .ColumnHeaders.Add , , "NIÑOS", 1000
    .HideColumnHeaders = False
    .FullRowSelect = True
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
End With
End Sub

Private Sub cmdAgregar_Click()
AgregaZona
End Sub

Private Sub cmdQuitar_Click()
If Me.lvZonas.ListItems.count = 0 Then Exit Sub
Me.lvZonas.ListItems.Remove Me.lvZonas.SelectedItem.Index
Me.cmdQuitar.Enabled = False
End Sub

Private Sub DatZonas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub DatZonas_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.txtAdultos.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
Dim oRsZonas As ADODB.Recordset
ConfiguraLV
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SpListarZonas"
Set oRsZonas = oCmdEjec.Execute(, LK_CODCIA)

With Me.DatZonas
Set .RowSource = oRsZonas
.ListField = "DENOMINA"
.BoundColumn = "CODIGO"
End With

 Dim Item As Object
 If orsMarcados Is Nothing Then Exit Sub
    If orsMarcados.RecordCount = 0 Then Exit Sub
    orsMarcados.MoveFirst
        Dim c As Integer
            Do While Not orsMarcados.EOF
            Set Item = Me.lvZonas.ListItems.Add(, , orsMarcados!ZONA)
            Item.Tag = orsMarcados!codzona
                Item.SubItems(1) = orsMarcados!adultos
                Item.SubItems(2) = orsMarcados!ninios
               
            
            orsMarcados.MoveNext
            Loop
        
        
End Sub

Private Sub lvZonas_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.cmdQuitar.Enabled = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index

        Case 1

            If Me.lvZonas.ListItems.count = 0 Then
                MsgBox "Debe agregar Al Menos 1 (UNA) Zona para el Contrato.", vbInformation, TituloSistema

                Exit Sub

            End If
            
            'frmContratos.oRsZonas.Close
            'frmContratos.oRsZonas.Fields.Delete 1
            If frmContratos.oRsZonas.RecordCount <> 0 Then
                frmContratos.oRsZonas.Filter = ""
                frmContratos.oRsZonas.MoveFirst

                'If frmContratos.oRsZonas.EOF Or frmContratos.oRsZonas.BOF Then
                While Not frmContratos.oRsZonas.EOF

                    frmContratos.oRsZonas.Delete
                    frmContratos.oRsZonas.MoveNext

                Wend

            End If

            Dim Item  As Object

            Dim VZONA As String, vADULTOS As Double, vNINIOS As Double

            vADULTOS = 0
            vNINIOS = 0

            For Each Item In Me.lvZonas.ListItems
 
                frmContratos.oRsZonas.AddNew
                frmContratos.oRsZonas!codzona = Item.Tag ' Me.lvZonas.ListItems(c).Tag
                frmContratos.oRsZonas!ZONA = Item.Text ' Me.lvZonas.ListItems(c).Text
                frmContratos.oRsZonas!adultos = Item.SubItems(1) ' Me.lvZonas.ListItems(c).SubItems(1)
                frmContratos.oRsZonas!ninios = Item.SubItems(2)
                frmContratos.oRsZonas.Update
                vADULTOS = vADULTOS + Item.SubItems(1)
                vNINIOS = vNINIOS + Item.SubItems(2)
                VZONA = VZONA + Item.Text + " - "
                ' End If

            Next

            frmContratos.oRsZonas.MoveFirst
            frmContratos.lblZonas.Caption = Left(VZONA, Len(Trim(VZONA)) - 2)
            frmContratos.txtAdultos.Text = vADULTOS
            frmContratos.txtNinios.Text = vNINIOS
            Unload Me

        Case 2
            Unload Me

    End Select

End Sub

Private Sub txtADULTOS_KeyPress(KeyAscii As Integer)
   If SoloNumeros(KeyAscii) Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Me.txtNinios.SetFocus
End Sub

Private Sub AgregaZona()

    If Me.DatZonas.BoundText = "" Then
        MsgBox "Debe elegir la zona para agregar.", vbInformation, TituloSistema
        Me.DatZonas.SetFocus
    ElseIf Len(Trim(Me.txtAdultos.Text)) = 0 Then
        MsgBox "Debe ingresar la cantidad de adultos.", vbCritical, TituloSistema
        Me.txtAdultos.SetFocus
    ElseIf val(Me.txtAdultos.Text) <= 0 Then
        MsgBox "La Cantidad de Adultos es incorrecta.", vbInformation, TituloSistema
        Me.txtAdultos.SelStart = 0
        Me.txtAdultos.SelLength = Len(Me.txtAdultos.Text)
        Me.txtAdultos.SetFocus
    Else

        Dim Item  As Object

        Dim vPasa As Boolean
        Dim vCapacidad As Integer
        
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SPDEVUELVECAPACIDADxZONA"
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adDouble, adParamInput, , Me.DatZonas.BoundText)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CAPACIDAD", adDouble, adParamOutput, , 0)
oCmdEjec.Execute

vCapacidad = oCmdEjec.Parameters("@CAPACIDAD").Value

If (val(Me.txtAdultos.Text) + IIf(Len(Trim(Me.txtNinios.Text)) = 0, 0, val(Me.txtNinios.Text))) > vCapacidad Then
MsgBox "Las Cantidades superan la Capacidad del Local.", vbInformation, TituloSistema
  Me.txtAdultos.SelStart = 0
        Me.txtAdultos.SelLength = Len(Me.txtAdultos.Text)
Me.txtAdultos.SetFocus
Exit Sub
End If

        If Me.lvZonas.ListItems.count = 0 Then 'agrega
            Set Item = Me.lvZonas.ListItems.Add(, , Me.DatZonas.Text)
            Item.Tag = Me.DatZonas.BoundText
            Item.SubItems(1) = Me.txtAdultos.Text
            Item.SubItems(2) = IIf(Len(Trim(Me.txtNinios.Text)) = 0, 0, Me.txtNinios.Text)
            Me.DatZonas.BoundText = ""
            Me.txtAdultos.Text = ""
            Me.txtNinios.Text = ""
            Me.DatZonas.SetFocus
        Else
    
            vPasa = True

            For Each Item In Me.lvZonas.ListItems
                If Item.Tag = Me.DatZonas.BoundText Then
                    vPasa = False
                    Exit For
                End If
            Next
    
            If vPasa Then
                Set Item = Me.lvZonas.ListItems.Add(, , Me.DatZonas.Text)
                Item.Tag = Me.DatZonas.BoundText
                Item.SubItems(1) = Me.txtAdultos.Text
                Item.SubItems(2) = IIf(Len(Trim(Me.txtNinios.Text)) = 0, 0, Me.txtNinios.Text)
                Me.DatZonas.BoundText = ""
                Me.txtAdultos.Text = ""
                Me.txtNinios.Text = ""
                Me.DatZonas.SetFocus
            Else
                MsgBox "La Zona ya se encuentra en la Lista.", vbCritical, TituloSistema
            End If
        End If
End If
    End Sub

Private Sub txtNINIOS_KeyPress(KeyAscii As Integer)
 If SoloNumeros(KeyAscii) Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then cmdAgregar_Click
End Sub
