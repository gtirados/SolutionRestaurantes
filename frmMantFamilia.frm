VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMantFamilia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Familia"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   8430
   Begin TabDlg.SSTab stabFamilia 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Familia"
      TabPicture(0)   =   "frmMantFamilia.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(4)=   "lblCodigo"
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(8)=   "txtDenominacion"
      Tab(0).Control(9)=   "txtImpresora"
      Tab(0).Control(10)=   "txtGrupo"
      Tab(0).Control(11)=   "txtDscto"
      Tab(0).Control(12)=   "ComVisible"
      Tab(0).Control(13)=   "txtImpresora2"
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Listado"
      TabPicture(1)   =   "frmMantFamilia.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtSearch"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lvFamilias"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.TextBox txtImpresora2 
         Height          =   285
         Left            =   -72240
         TabIndex        =   17
         Top             =   2160
         Width           =   3375
      End
      Begin VB.ComboBox ComVisible 
         Height          =   315
         ItemData        =   "frmMantFamilia.frx":0038
         Left            =   -72240
         List            =   "frmMantFamilia.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox txtDscto 
         Height          =   285
         Left            =   -72240
         TabIndex        =   13
         Tag             =   "X"
         Top             =   3120
         Width           =   1335
      End
      Begin MSComctlLib.ListView lvFamilias 
         Height          =   3495
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   6165
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
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   480
         Width           =   7095
      End
      Begin VB.TextBox txtGrupo 
         Height          =   285
         Left            =   -72240
         MaxLength       =   1
         TabIndex        =   9
         Tag             =   "X"
         Top             =   2640
         Width           =   3375
      End
      Begin VB.TextBox txtImpresora 
         Height          =   285
         Left            =   -72240
         TabIndex        =   7
         Tag             =   "X"
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox txtDenominacion 
         Height          =   285
         Left            =   -72240
         TabIndex        =   6
         Tag             =   "X"
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Impresora2:"
         Height          =   195
         Left            =   -73485
         TabIndex        =   18
         Top             =   2160
         Width           =   1080
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visible:"
         Height          =   195
         Left            =   -73035
         TabIndex        =   16
         Top             =   3660
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Dscto:"
         Height          =   195
         Left            =   -73200
         TabIndex        =   15
         Top             =   3165
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Familia:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   675
      End
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -72240
         TabIndex        =   8
         Tag             =   "X"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo:"
         Height          =   195
         Left            =   -73005
         TabIndex        =   5
         Top             =   2640
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Denominación:"
         Height          =   195
         Left            =   -73695
         TabIndex        =   4
         Top             =   1200
         Width           =   1290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Impresora:"
         Height          =   195
         Left            =   -73380
         TabIndex        =   3
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   195
         Left            =   -73080
         TabIndex        =   2
         Top             =   720
         Width           =   675
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8400
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantFamilia.frx":004E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantFamilia.frx":03E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantFamilia.frx":0782
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantFamilia.frx":0B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantFamilia.frx":0EB6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbFamilia 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   635
      ButtonWidth     =   1826
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Guardar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Editar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eliminar"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMantFamilia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VNuevo As Boolean
Private Sub Estado_Botones(val As Valores)
Select Case val
    Case InicializarFormulario, grabar, cancelar, Eliminar
        Me.tbFamilia.Buttons(1).Enabled = True
        Me.tbFamilia.Buttons(2).Enabled = False
        Me.tbFamilia.Buttons(3).Enabled = False
        Me.tbFamilia.Buttons(4).Enabled = False
        Me.tbFamilia.Buttons(5).Enabled = False
        Me.stabFamilia.tab = 1
    Case Nuevo, Editar
        Me.tbFamilia.Buttons(1).Enabled = False
        Me.tbFamilia.Buttons(2).Enabled = True
        Me.tbFamilia.Buttons(3).Enabled = False
        Me.tbFamilia.Buttons(4).Enabled = True
        Me.lvFamilias.Enabled = False
        Me.txtSearch.Enabled = False
        Me.stabFamilia.tab = 0
        Me.tbFamilia.Buttons(5).Enabled = False
    Case buscar
        Me.tbFamilia.Buttons(1).Enabled = True
        Me.tbFamilia.Buttons(2).Enabled = False
        Me.tbFamilia.Buttons(3).Enabled = False
        Me.tbFamilia.Buttons(4).Enabled = False
        Me.stabFamilia.tab = 1
    Case AntesDeActualizar
        Me.tbFamilia.Buttons(1).Enabled = False
        Me.tbFamilia.Buttons(2).Enabled = False
        Me.tbFamilia.Buttons(3).Enabled = True
        Me.tbFamilia.Buttons(4).Enabled = True
         'Me.tbFamilia.Buttons(5).Enabled = True
        Me.stabFamilia.tab = 0
End Select
End Sub
Private Sub ConfigurarLV()
With Me.lvFamilias
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "Familia", 3000
    .ColumnHeaders.Add , , "Impresora", 1500
    .ColumnHeaders.Add , , "Grupo", 1400
    .ColumnHeaders.Add , , "Dscto", 0
    .ColumnHeaders.Add , , "Visible", 600
    .ColumnHeaders.Add , , "PRINT2", 0
End With
End Sub


Private Sub Form_Load()
ConfigurarLV
DesactivarControles Me
Estado_Botones InicializarFormulario
RealizarBusqueda
CentrarFormulario MDIForm1, Me
End Sub

Private Sub RealizarBusqueda(Optional vSearch As String = "")
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_FAMILIA_SEARCH"
    Me.lvFamilias.ListItems.Clear

    Dim ORSf As ADODB.Recordset

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

    If Len(Trim(vSearch)) <> 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SEARCH", adVarChar, adParamInput, 50, Trim(Me.txtSearch.Text))
    End If
    
    Set ORSf = oCmdEjec.Execute
    
    Do While Not ORSf.EOF

        With Me.lvFamilias.ListItems.Add(, , Trim(ORSf!nom))
            .Tag = Trim(ORSf!IDE)
            .SubItems(1) = ORSf!Print
            .SubItems(2) = ORSf!grupo
            .SubItems(3) = ORSf!dscto
        .SubItems(4) = ORSf!Visible
        .SubItems(5) = ORSf!PRINT2
        End With
   
        ORSf.MoveNext
    Loop

End Sub

Private Sub lvFamilias_Click()
Me.tbFamilia.Buttons(5).Enabled = True
End Sub

Private Sub lvFamilias_DblClick()
If Me.lvFamilias.ListItems.count <> 0 Then Mandar_Datos
End Sub



Private Sub tbFamilia_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index

        Case 1 'NUEVO
            ActivarControles Me
            LimpiarControles Me
            Estado_Botones Nuevo
            VNuevo = True
            Me.txtDenominacion.SetFocus

        Case 2 'GUARDAR
            LimpiaParametros oCmdEjec

            If Len(Trim(Me.txtDenominacion.Text)) = 0 Then
                MsgBox "Debe ingresar la Denominación de la Mesa", vbCritical, Pub_Titulo
                Me.txtDenominacion.SetFocus
                '            ElseIf Len(Trim(Me.txtImpresora.Text)) = 0 Then
                '                MsgBox "Debe ingresar la impresora", vbCritical, Pub_Titulo
                '                Me.txtImpresora.SetFocus
                '            ElseIf Len(Trim(Me.txtGrupo.Text)) = 0 Then
                '                MsgBox "Debe ingresar el Grupo.", vbCritical, Pub_Titulo
                '                Me.txtGrupo.SetFocus
            Else

                On Error GoTo grabar
        
                LimpiaParametros oCmdEjec
                oCmdEjec.Prepared = True
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DENOMINACION", adVarChar, adParamInput, 50, Trim(Me.txtDenominacion.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IMPRESORA", adVarChar, adParamInput, 50, Trim(Me.txtImpresora.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IMPRESORA2", adVarChar, adParamInput, 50, Trim(Me.txtImpresora2.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@GRUPO", adChar, adParamInput, 1, Me.txtGrupo.Text)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DSCTO", adDouble, adParamInput, , Me.txtDscto.Text)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@VISIBLE", adBoolean, adParamInput, , Me.ComVisible.ListIndex)

                If VNuevo Then
                    oCmdEjec.CommandText = "SP_FAMILIA_REGISTRAR"
                Else
                    oCmdEjec.CommandText = "SP_FAMILIA_MODIFICAR"
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODIGO", adInteger, adParamInput, , Me.lblCodigo.Caption)
                End If

                oCmdEjec.Execute
                'oCmdEjec.Execute , Array(CodCia, Trim(Me.txtCodigo.Text), Trim(Me.txtDenominacion.Text), Trim(Me.txtZona.Text))
                DesactivarControles Me
                Estado_Botones grabar
                
                Me.lvFamilias.Enabled = True
                Me.txtSearch.Enabled = True

                If VNuevo Then
        
                    '                    With Me.lvFamilias.ListItems.Add(, , Trim(Me.txtDenominacion.Text))
                    '                        .Tag = Trim(Me.lblCodigo.Caption)
                    '                        .SubItems(1) = Me.txtImpresora.Text
                    '                        .SubItems(2) = Me.txtGrupo.Text
                    '                    End With
                    RealizarBusqueda Me.txtSearch.Text
            
                Else
                    Me.lvFamilias.SelectedItem.Text = Trim(Me.txtDenominacion.Text)
                    Me.lvFamilias.SelectedItem.SubItems(1) = Me.txtImpresora.Text
                    Me.lvFamilias.SelectedItem.SubItems(2) = Me.txtGrupo.Text
                    Me.lvFamilias.SelectedItem.SubItems(3) = Me.txtDscto.Text

                    If Me.ComVisible.ListIndex = 0 Then
                        Me.lvFamilias.SelectedItem.SubItems(4) = "NO"
                    Else
                        Me.lvFamilias.SelectedItem.SubItems(4) = "SI"
                    End If
                    Me.lvFamilias.SelectedItem.SubItems(5) = Me.txtImpresora2.Text
                End If
        
                'set itemg=me.lvMesas.ListItems.Add(,,
                MsgBox "Datos Almacenados Correctamente", vbInformation, Pub_Titulo

                Exit Sub

grabar:
                MsgBox Err.Description, vbInformation, Pub_Titulo

            End If

        Case 3 'MODIFICAR
            VNuevo = False
            Estado_Botones Editar
            ActivarControles Me
            'Me.txtCodigo.Enabled = False
            Me.txtDenominacion.SetFocus

        Case 4 'CANCELAR
            Estado_Botones cancelar
            DesactivarControles Me
            Me.lvFamilias.Enabled = True
            Me.txtSearch.Enabled = True
            Me.lvFamilias.SelectedItem.Selected = False
            
        Case 5 'ELIMINAR

            On Error GoTo elimina

            If MsgBox("¿Desea continuar con la operación.?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
                LimpiaParametros oCmdEjec
                oCmdEjec.CommandText = "SP_FAMILIA_ELIMINAR"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODIGO", adBigInt, adParamInput, , Me.lvFamilias.SelectedItem.Tag)
                oCmdEjec.Execute
                
                Me.lvFamilias.ListItems.Remove Me.lvFamilias.SelectedItem.Index
                Me.tbFamilia.Buttons(5).Enabled = False
            
                MsgBox "Datos Eliminados Correctamente", vbInformation, Pub_Titulo
            End If

            Exit Sub

elimina:
            MsgBox Err.Description, vbCritical, Pub_Titulo
            
    End Select

End Sub

Private Sub txtDscto_KeyPress(KeyAscii As Integer)
  If NumerosyPunto(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtSearch_Change()
RealizarBusqueda Me.txtSearch.Text
End Sub

Sub Mandar_Datos()

    With Me.lvFamilias
        Me.lblCodigo.Caption = .SelectedItem.Tag
        Me.txtDenominacion.Text = .SelectedItem.Text
        'Me.txtDenominacion.Text = Trim(.SelectedItem.SubItems(1))
        'Me.txtZona.Text = Trim(.SelectedItem.SubItems(2))
        Me.txtImpresora.Text = .SelectedItem.SubItems(1)
        Me.txtGrupo.Text = .SelectedItem.SubItems(2)
        Me.txtDscto.Text = .SelectedItem.SubItems(3)
        If .SelectedItem.SubItems(4) = "SI" Then
            Me.ComVisible.ListIndex = 1
        Else
        Me.ComVisible.ListIndex = 0
        End If
        Me.txtImpresora2.Text = .SelectedItem.SubItems(5)
        Estado_Botones AntesDeActualizar
    End With

End Sub
