VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMantSubFamilia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Sub Familias"
   ClientHeight    =   4980
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   8430
   Begin TabDlg.SSTab stabFamilia 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Sub Familia"
      TabPicture(0)   =   "frmMantSubFamilia.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCodigo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtDenominacion"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DatFamilia"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Listado"
      TabPicture(1)   =   "frmMantSubFamilia.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "datFamiliasearch"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtSearch"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lvSubFamilia"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label4"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin MSDataListLib.DataCombo datFamiliasearch 
         Height          =   315
         Left            =   -73560
         TabIndex        =   13
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   -73560
         TabIndex        =   10
         Top             =   360
         Width           =   6615
      End
      Begin MSDataListLib.DataCombo DatFamilia 
         Height          =   315
         Left            =   2760
         TabIndex        =   8
         Top             =   2280
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtDenominacion 
         Height          =   285
         Left            =   2760
         TabIndex        =   1
         Tag             =   "X"
         Top             =   1680
         Width           =   3375
      End
      Begin MSComctlLib.ListView lvSubFamilia 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   9
         Top             =   1080
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   5530
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Familia:"
         Height          =   195
         Left            =   -74400
         TabIndex        =   12
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SubFamilia:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   11
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   195
         Left            =   1935
         TabIndex        =   6
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Familia:"
         Height          =   195
         Left            =   2040
         TabIndex        =   5
         Top             =   2280
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Denominación:"
         Height          =   195
         Left            =   1320
         TabIndex        =   4
         Top             =   1680
         Width           =   1290
      End
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2760
         TabIndex        =   3
         Tag             =   "X"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Familia:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   2
         Top             =   480
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
            Picture         =   "frmMantSubFamilia.frx":0038
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantSubFamilia.frx":03D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantSubFamilia.frx":076C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantSubFamilia.frx":0B06
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantSubFamilia.frx":0EA0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbFamilia 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8430
      _ExtentX        =   14870
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
Attribute VB_Name = "frmMantSubFamilia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VNuevo As Boolean

Private Sub datFamiliasearch_Change()
RealizarBusqueda
End Sub

Private Sub Form_Load()
ConfigurarLV
DesactivarControles Me
Estado_Botones InicializarFormulario

CentrarFormulario MDIForm1, Me
Me.datFamiliasearch.Enabled = True
LlenarFamilias
RealizarBusqueda
End Sub

Private Sub LlenarFamilias()
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SP_SUBFAMILIA_FAMILIA_LIST"

Dim orsFS As ADODB.Recordset
    Set orsFS = oCmdEjec.Execute(, LK_CODCIA)
    Set Me.datFamiliasearch.RowSource = orsFS
    Me.datFamiliasearch.ListField = orsFS.Fields(1).Name
    Me.datFamiliasearch.BoundColumn = orsFS.Fields(0).Name
    Me.datFamiliasearch.BoundText = -1
End Sub

Private Sub ConfigurarLV()
With Me.lvSubFamilia
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "Familia", 3000
    .ColumnHeaders.Add , , "Sub Familia", 3000
    .ColumnHeaders.Add , , "IDEFamilia", 0
    
    
End With
End Sub

Private Sub Estado_Botones(val As Valores)
Select Case val
    Case InicializarFormulario, grabar, cancelar, Eliminar
        Me.tbFamilia.Buttons(1).Enabled = True
        Me.tbFamilia.Buttons(2).Enabled = False
        Me.tbFamilia.Buttons(3).Enabled = False
        Me.tbFamilia.Buttons(4).Enabled = False
        'Me.tbFamilia.Buttons(5).Enabled = False
        Me.tbFamilia.Buttons(5).Enabled = False
        Me.stabFamilia.tab = 1
    Case Nuevo, Editar
        Me.tbFamilia.Buttons(1).Enabled = False
        Me.tbFamilia.Buttons(2).Enabled = True
        Me.tbFamilia.Buttons(3).Enabled = False
        Me.tbFamilia.Buttons(4).Enabled = True
       '  Me.tbFamilia.Buttons(5).Enabled = False
        Me.lvSubFamilia.Enabled = False
        Me.txtSearch.Enabled = False
        Me.datFamiliasearch.Enabled = False
        Me.tbFamilia.Buttons(5).Enabled = False
        Me.stabFamilia.tab = 0
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

Private Sub RealizarBusqueda(Optional vSearch As String = "")
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_SUBFAMILIA_SEARCH"
    Me.lvSubFamilia.ListItems.Clear

    Dim ORSf As ADODB.Recordset

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDFAMILIA", adBigInt, adParamInput, , Me.datFamiliasearch.BoundText)
    
    If Len(Trim(vSearch)) <> 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SEARCH", adVarChar, adParamInput, 50, Trim(Me.txtSearch.Text))
    End If
    
    Set ORSf = oCmdEjec.Execute
    
    Do While Not ORSf.EOF

        With Me.lvSubFamilia.ListItems.Add(, , Trim(ORSf!nom))
            .Tag = Trim(ORSf!IDE)
            .SubItems(1) = Trim(ORSf!Familia)
            .SubItems(2) = ORSf!IDEFAMILIA
        
        End With
   
        ORSf.MoveNext
    Loop

End Sub

Private Sub lvSubFamilia_Click()
Me.tbFamilia.Buttons(5).Enabled = True
End Sub

Private Sub lvSubFamilia_DblClick()
If Me.lvSubFamilia.ListItems.count <> 0 Then Mandar_Datos
End Sub

Private Sub tbFamilia_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index

        Case 1 'NUEVO
            ActivarControles Me
            LimpiarControles Me
            Estado_Botones Nuevo
            VNuevo = True
            CargarCombo
            Me.txtDenominacion.SetFocus

        Case 2 'GUARDAR
            LimpiaParametros oCmdEjec

            If Len(Trim(Me.txtDenominacion.Text)) = 0 Then
                MsgBox "Debe ingresar la Denominación de la SubFamilia", vbCritical, Pub_Titulo
                Me.txtDenominacion.SetFocus
            ElseIf Me.DatFamilia.BoundText = "-1" Then
                MsgBox "Debe elgir la Familia.", vbCritical, Pub_Titulo
                Me.DatFamilia.SetFocus
            Else

                On Error GoTo grabar
        
                LimpiaParametros oCmdEjec
                oCmdEjec.Prepared = True
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DENOMINACION", adVarChar, adParamInput, 50, Trim(Me.txtDenominacion.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDFAMILIA", adBigInt, adParamInput, , Me.DatFamilia.BoundText)

                If VNuevo Then
                    oCmdEjec.CommandText = "SP_SUBFAMILIA_REGISTRAR"
                Else
                    oCmdEjec.CommandText = "SP_SUBFAMILIA_MODIFICAR"
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODIGO", adInteger, adParamInput, , Me.lblCodigo.Caption)
                End If

                oCmdEjec.Execute
                'oCmdEjec.Execute , Array(CodCia, Trim(Me.txtCodigo.Text), Trim(Me.txtDenominacion.Text), Trim(Me.txtZona.Text))
                DesactivarControles Me
                Estado_Botones grabar
                
                Me.lvSubFamilia.Enabled = True
                Me.txtSearch.Enabled = True

                If VNuevo Then
        
                    With Me.lvSubFamilia.ListItems.Add(, , Trim(Me.txtDenominacion.Text))
                        .Tag = Trim(Me.lblCodigo.Caption)
                        .SubItems(1) = Me.DatFamilia.Text
                        .SubItems(2) = Me.DatFamilia.BoundText
                    End With
            
                Else
                    Me.lvSubFamilia.SelectedItem.Text = Trim(Me.txtDenominacion.Text)
                    Me.lvSubFamilia.SelectedItem.SubItems(1) = Me.DatFamilia.Text
                    Me.lvSubFamilia.SelectedItem.SubItems(2) = Me.DatFamilia.BoundText
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
            Me.lvSubFamilia.Enabled = True
            Me.txtSearch.Enabled = True
            Me.datFamiliasearch.Enabled = True
            Me.lvSubFamilia.SelectedItem.Selected = False

        Case 5 'ELIMINAR

            On Error GoTo elimina

            If MsgBox("¿Desea continuar con la operación.?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
                LimpiaParametros oCmdEjec
                oCmdEjec.CommandText = "SP_SUBFAMILIA_ELIMINAR"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODIGO", adBigInt, adParamInput, , Me.lvSubFamilia.SelectedItem.Tag)
                oCmdEjec.Execute
                
                Me.lvSubFamilia.ListItems.Remove Me.lvSubFamilia.SelectedItem.Index
                Me.tbFamilia.Buttons(5).Enabled = False
            
                MsgBox "Datos Eliminados Correctamente", vbInformation, Pub_Titulo
            End If

            Exit Sub

elimina:
            MsgBox Err.Description, vbCritical, Pub_Titulo
    End Select

End Sub

Private Sub txtSearch_Change()
RealizarBusqueda Me.txtSearch.Text
End Sub

Sub Mandar_Datos()
    CargarCombo

    With Me.lvSubFamilia
        Me.lblCodigo.Caption = .SelectedItem.Tag
        Me.txtDenominacion.Text = .SelectedItem.Text
        'Me.txtDenominacion.Text = Trim(.SelectedItem.SubItems(1))
        'Me.txtZona.Text = Trim(.SelectedItem.SubItems(2))
        Me.DatFamilia.BoundText = .SelectedItem.SubItems(2)
        Me.tbFamilia.Buttons(5).Enabled = False
        Estado_Botones AntesDeActualizar
    End With

End Sub

Private Sub CargarCombo()
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SP_FAMILIA_LIST"

Dim ORSd As ADODB.Recordset
Set ORSd = oCmdEjec.Execute(, LK_CODCIA)

Set Me.DatFamilia.RowSource = ORSd
    Me.DatFamilia.BoundColumn = ORSd.Fields(0).Name
    Me.DatFamilia.ListField = ORSd.Fields(1).Name
    Me.DatFamilia.BoundText = "-1"
    
End Sub
