VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConversion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conversión de Materias Primas"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConversion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   8415
   Begin VB.TextBox txtNro 
      Height          =   285
      Left            =   6840
      TabIndex        =   21
      Top             =   480
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2055
      Left            =   2160
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   740
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3625
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   635
      ButtonWidth     =   2196
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Procesar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7680
      Top             =   4200
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
            Picture         =   "frmConversion.frx":0D42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConversion.frx":1A94
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdEdit 
      Enabled         =   0   'False
      Height          =   360
      Left            =   7680
      Picture         =   "frmConversion.frx":1E2E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   8175
      Begin VB.TextBox txtCantidad 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CANTIDAD A PROCESAR:"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   120
         Width           =   2220
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dato Seleccionado"
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   8175
      Begin VB.Label lblStock 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5880
         TabIndex        =   20
         Top             =   960
         Width           =   1875
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STOCK:"
         Height          =   195
         Left            =   5040
         TabIndex        =   19
         Top             =   960
         Width           =   690
      End
      Begin VB.Label lblUnidad 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5880
         TabIndex        =   14
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label lblCosto 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2040
         TabIndex        =   13
         Top             =   960
         Width           =   1635
      End
      Begin VB.Label lblMP 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2040
         TabIndex        =   12
         Top             =   600
         Width           =   5715
      End
      Begin VB.Label lblCodigo 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "UNIDAD:"
         Height          =   195
         Left            =   4920
         TabIndex        =   10
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "COSTO:"
         Height          =   195
         Left            =   1080
         TabIndex        =   9
         Top             =   1020
         Width           =   705
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "MATERIA PRIMA:"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   660
         Width           =   1485
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CODIGO:"
         Height          =   195
         Left            =   960
         TabIndex        =   7
         Top             =   300
         Width           =   825
      End
   End
   Begin VB.TextBox txtMP 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
   Begin MSComctlLib.ListView lvSMP 
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   4895
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblTotalCant 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   6000
      TabIndex        =   17
      Top             =   6000
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "MATERIAS PRIMAS:"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   1710
   End
End
Attribute VB_Name = "frmConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private loc_key  As Integer
Private Sub ConfigurarLV()
With Me.lvSMP
    
    .ColumnHeaders.Add , , "MATERIA PRIMA", 3000
    .ColumnHeaders.Add , , "PRECIO"
    .ColumnHeaders.Add , , "UNIDAD"
    .ColumnHeaders.Add , , "CANTIDAD"
    .LabelEdit = lvwManual
    .FullRowSelect = True
    .Gridlines = True
    .View = lvwReport
End With

With Me.ListView1
    .ColumnHeaders.Add , , "CODIGO", 700
    .ColumnHeaders.Add , , "MATERIA PRIMA"
    .ColumnHeaders.Add , , "COSTO"
    .ColumnHeaders.Add , , "STOCK"
    .ColumnHeaders.Add , , "UNIDAD"
    .FullRowSelect = True
    .Gridlines = True
    .HideColumnHeaders = True
    .View = lvwReport
    .HideSelection = False
End With
End Sub

Private Sub cmdEdit_Click()

    If Len(Trim(Me.txtCantidad.Text)) = 0 Then
        MsgBox "Debe ingresar la Cantidad a Procesar.", vbInformation, Pub_Titulo
        Me.txtCantidad.SetFocus

        Exit Sub

    End If

    If Not IsNumeric(Me.txtCantidad.Text) Then
        MsgBox "La Cantidad Proporcionada es incorrecta.", vbCritical, Pub_Titulo
        Me.txtCantidad.SetFocus
        Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)

        Exit Sub

    End If

    If val(Me.txtCantidad.Text) = 0 Then
        MsgBox "La Cantidad Proporcionada es incorrecta.", vbCritical, Pub_Titulo
        Me.txtCantidad.SetFocus
        Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)

        Exit Sub

    End If
    
    If val(Me.txtCantidad.Text) > val(Me.lblStock.Caption) Then
        MsgBox "La Cantidad a Procesar es superior al Stock Actual.", vbCritical, Pub_Titulo
        Me.txtCantidad.SetFocus
        Me.txtCantidad.SelStart = 0
        Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)

        Exit Sub
    End If

Dim C As Integer
Dim SUM As Double
SUM = 0
For C = 1 To Me.lvSMP.ListItems.count
SUM = SUM + IIf(Len(Trim(Me.lvSMP.ListItems(C).SubItems(3))) = 0, 0, Me.lvSMP.ListItems(C).SubItems(3))
Next
frmConversionEdit.txtCantidad.Text = frmConversion.txtCantidad.Text - SUM

    frmConversionEdit.Show vbModal

    If frmConversionEdit.vCant <> 0 Then
        Me.lvSMP.SelectedItem.SubItems(3) = frmConversionEdit.vCant
        Dim itemB As Object
        Dim VCAT As Double
        VCAT = 0
        For Each itemB In Me.lvSMP.ListItems
            VCAT = VCAT + IIf(Len(Trim(itemB.SubItems(3))) = 0, 0, itemB.SubItems(3))
        Next
        Me.lblTotalCant.Caption = str(VCAT)
    End If

End Sub

Private Sub Form_Load()
ConfigurarLV
CentrarFormulario MDIForm1, Me
Me.lvSMP.Enabled = False
Me.Toolbar1.Buttons(3).Enabled = False
Me.Toolbar1.Buttons(2).Enabled = False
Me.Toolbar1.Buttons(4).Enabled = False
End Sub

Private Sub lvSMP_DblClick()
cmdEdit_Click
End Sub

Private Sub lvSMP_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.cmdEdit.Enabled = True
End Sub

Private Sub lvSMP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
If Me.lvSMP.SelectedItem Is Nothing Then Exit Sub
    Me.lvSMP.ListItems.Remove Me.lvSMP.SelectedItem.Index
End If
End Sub

Private Sub lvSMP_KeyPress(KeyAscii As Integer)
cmdEdit_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
Case 1
Limpia
Me.txtNro.Enabled = False
Me.txtMP.Enabled = True
Me.Toolbar1.Buttons(1).Enabled = False
Me.Toolbar1.Buttons(2).Enabled = True
Me.Toolbar1.Buttons(3).Enabled = True
Me.txtMP.SetFocus
        Case 2
        
            If Len(Trim(Me.lblCodigo.Caption)) = 0 Then
                MsgBox "Debe elegir un Insumo.", vbInformation, Pub_Titulo
                Me.txtMP.SetFocus

                Exit Sub

            End If

            If Len(Trim(Me.lblTotalCant.Caption)) = 0 Or Len(Trim(Me.lblTotalCant.Caption)) = 0 Then
                MsgBox "Debe Agregar Cantidades a las Sub Materias Primas.", vbCritical, Pub_Titulo

                Exit Sub

            End If
    
            If Len(Trim(Me.txtCantidad.Text)) = 0 Then
                MsgBox "Debe ingrear la Cantidad a Procesar.", vbInformation, Pub_Titulo
                Me.txtCantidad.SetFocus

                Exit Sub

            End If
    
            If val(Me.txtCantidad.Text) < val(Me.lblTotalCant.Caption) Then
                MsgBox "La Cantidad  a Procesar debe ser Mayor o igual" & vbCrLf & "a la cantidad de Sub Materias primas.", vbCritical, Pub_Titulo
                Me.txtCantidad.SetFocus

                Exit Sub

            End If
    
            If Not IsNumeric(Me.txtCantidad.Text) Then
                MsgBox "La Cantidad proporcionada es incorrecta.", vbInformation, Pub_Titulo
                Me.txtCantidad.SetFocus
                Me.txtCantidad.SelStart = 0
                Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)

                Exit Sub

            End If
            
            If val(Me.lblTotalCant.Caption) < Me.txtCantidad.Text Then
                MsgBox "El Total de SUb Materias primas debe ser igual que la cantidad a procesar.", vbInformation, Pub_Titulo
                Me.txtCantidad.SetFocus
                Me.txtCantidad.SelStart = 0
                Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)

                Exit Sub

            End If
    
            Dim STRSMP As String
    
            Dim iO     As Object

            STRSMP = "<r>"

            For Each iO In Me.lvSMP.ListItems

                STRSMP = STRSMP & "<d "
                STRSMP = STRSMP & "cp=""" & iO.Tag & """ "
                STRSMP = STRSMP & "pm=""" & iO.SubItems(1) & """ "
                STRSMP = STRSMP & "ct=""" & iO.SubItems(3) & """ "
                STRSMP = STRSMP & "un=""" & iO.SubItems(2) & """ "
                STRSMP = STRSMP & "/>"
            Next

            STRSMP = STRSMP & "</r>"
            
            On Error GoTo grabar:
            Dim Vnro As Integer
            Pub_ConnAdo.BeginTrans
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SP_PROCESAR_MATERIAPRIMA"
            
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , Me.lblCodigo.Caption)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
                
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CANTIDAD", adDouble, adParamInput, 2, Me.txtCantidad.Text)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, LK_CODUSU)
                
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CAMBIO", adDouble, adParamInput, , LK_TIPO_CAMBIO)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@UNIDAD", adVarChar, adParamInput, 15, Me.lblUnidad.Caption)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@costo", adDouble, adParamInput, , Me.lblCosto.Caption)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CANTPROC", adDouble, adParamInput, , Me.txtCantidad.Text)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SMP", adBSTR, adParamInput, 80000, STRSMP)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ALL_NUMFAC", adBigInt, adParamOutput, , Vnro)
            oCmdEjec.Execute
            Pub_ConnAdo.CommitTrans
            Me.txtNro.Text = oCmdEjec.Parameters("@ALL_NUMFAC").Value
            
            MsgBox "Datos Procesados Correctamente.", vbInformation, Pub_Titulo
            Me.Toolbar1.Buttons(3).Enabled = False
            Me.Toolbar1.Buttons(2).Enabled = False
            Me.Toolbar1.Buttons(1).Enabled = True
            Me.Toolbar1.Buttons(4).Enabled = True

            Exit Sub

grabar:
            Pub_ConnAdo.RollbackTrans
            MsgBox Err.Description, vbCritical, Pub_Titulo

        Case 3
            Limpia
            Me.lvSMP.Enabled = False
            Me.txtMP.Enabled = False
            Me.txtNro.Enabled = True
            
            Me.Toolbar1.Buttons(1).Enabled = True
            Me.Toolbar1.Buttons(2).Enabled = False
            Me.Toolbar1.Buttons(3).Enabled = False
        Case 4 'imprime
            If Len(Trim(Me.txtNro.Text)) = 0 Then
                MsgBox "Debe Ingresar el Nro para Imprimir.", vbInformation, Pub_Titulo
                Me.txtNro.SetFocus
                Exit Sub
            End If
            If Not IsNumeric(Me.txtNro.Text) Then
                MsgBox "El numero proporcionado es incorrecto.", vbCritical, Pub_Titulo
                Me.txtNro.SetFocus
                Exit Sub
            End If
          Imprime Me.txtNro.Text
    End Select

End Sub

Private Sub Imprime(xNro As Integer)
   On Error GoTo Ver
                Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
                
                Dim crParamDef  As CRAXDRT.ParameterFieldDefinition
                
                Dim objCrystal  As New CRAXDRT.APPLICATION
                
                Dim RutaReporte As String
                Dim oUSER As String, oCLAVE As String, oLOCAL As String
                
                
                 'DATOS COMPLEMENTARIOS
                Dim orsC As ADODB.Recordset
                LimpiaParametros oCmdEjec
                oCmdEjec.CommandText = "SPDATOSCOMPLEMENTARIOSCONTRATOS"
                        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                
                Set orsC = oCmdEjec.Execute
                
                
                'RutaReporte = "d:\VISTACONTRATO.rpt"
                RutaReporte = Trim(orsC!RutaReporte) + "CONVERSION.rpt"
                oUSER = orsC!USUARIO
                oCLAVE = orsC!Clave
                oLOCAL = orsC!LOCAL
                
                If VReporte Is Nothing Then VReporte = New CRAXDRT.Report
    
                Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
                Set crParamDefs = VReporte.ParameterFields
                
                For Each crParamDef In crParamDefs
                
                    Select Case crParamDef.ParameterFieldName
                
                        Case "TITULO"
                            crParamDef.AddCurrentValue "CONTRATO CENA " & oLOCAL & " - Nº - " & contextEvent.ScheduleID ' Me.lvData.SelectedItem.Text
                    End Select
                
                Next
                
                LimpiaParametros oCmdEjec
                oCmdEjec.CommandType = adCmdStoredProc
                oCmdEjec.CommandText = "SP_CONVERSION_PRINT"
                'oCmdEjec.CommandText = "SpPrintComanda"
                
                Dim rsd As ADODB.Recordset
                
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMERO", adBigInt, adParamInput, , xNro)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                
                Set rsd = oCmdEjec.Execute
           
               
                
                VReporte.DataBase.SetDataSource rsd, , 1  'lleno el objeto reporte
          
               
                frmContratosReporte.crContrato.ReportSource = VReporte
                
              
                frmContratosReporte.crContrato.ViewReport
                
                frmContratosReporte.Show
                Set VReporte = Nothing
                Set VReporteS = Nothing
                
                'RutaReporte = "C:\Admin\Nordi\Comanda1.rpt"
             Exit Sub
Ver:
             MsgBox Err.Description, vbCritical, Pub_Titulo
End Sub

Private Sub Limpia()
  Me.lvSMP.ListItems.Clear
            Me.cmdEdit.Enabled = False
            Me.lblCodigo.Caption = ""
            Me.lblUnidad.Caption = ""
            Me.lblCosto.Caption = ""
            Me.lblMP.Caption = ""
            Me.txtMP.Text = ""
            Me.lblStock.Caption = ""
            Me.txtCantidad.Text = ""
            Me.lblTotalCant.Caption = ""
End Sub
Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If SoloNumeros(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtMP_Change()
Dim ORSDatos As ADODB.Recordset

    Me.ListView1.ListItems.Clear
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPMATERIAPRIMA_LISTADO"
    Set ORSDatos = oCmdEjec.Execute(, Array(LK_CODCIA, Me.txtMP.Text))

    Dim itemX As Object
        
    If Not ORSDatos.EOF Then

        Do While Not ORSDatos.EOF
            Set itemX = Me.ListView1.ListItems.Add(, , Trim(ORSDatos!Codigo))
            itemX.SubItems(1) = Trim(ORSDatos.Fields(1).Value)
            itemX.SubItems(2) = ORSDatos!costo
            itemX.SubItems(3) = ORSDatos!stock
            itemX.SubItems(4) = Trim(ORSDatos!UNIDAD)
            ORSDatos.MoveNext
        Loop

        Me.ListView1.Visible = True
        Me.ListView1.ListItems(1).Selected = True
        loc_key = 1
        Me.ListView1.ListItems(1).EnsureVisible
        vBuscar = False
        '            Else
        '         If MsgBox("Cliente no existe." + vbCrLf + "¿Desea Crearlo.?", vbQuestion + vbYesNo, "Restaurantes") = vbYes Then
        '         frmCLI.Show vbModal
        '         End If
    Else
        loc_key = -1
    End If
End Sub

Private Sub txtMP_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error GoTo SALE

    Dim strFindMe As String

    Dim itmFound  As Object ' ListItem    ' Variable FoundItem.

    If KeyCode = 40 Then  ' flecha abajo
        loc_key = loc_key + 1

        If loc_key > Me.ListView1.ListItems.count Then loc_key = Me.ListView1.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 38 Then
        loc_key = loc_key - 1

        If loc_key < 1 Then loc_key = 1
        GoTo posicion
    End If

    If KeyCode = 34 Then
        loc_key = loc_key + 17

        If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 33 Then
        loc_key = loc_key - 17

        If loc_key < 1 Then loc_keyP = 1
        GoTo posicion
    End If

    If KeyCode = 27 Then
        Me.ListView1.Visible = False

    End If

    GoTo fin
posicion:
    ListView1.ListItems.Item(loc_key).Selected = True
    ListView1.ListItems.Item(loc_key).EnsureVisible
    'txtRS.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
    'Me.txtTransportista.SelStart = Len(txtTransportista.Text)
    Me.txtMP.SelStart = Len(Me.txtMP.Text)
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtMP_KeyPress(KeyAscii As Integer)
 KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then
    
    If loc_key <= 0 Then Exit Sub
    
     Me.lblCodigo.Caption = Me.ListView1.ListItems(loc_key).Text
     Me.lblMP.Caption = Me.ListView1.ListItems(loc_key).SubItems(1)
     Me.lblCosto.Caption = Me.ListView1.ListItems(loc_key).SubItems(2)
     Me.lblStock.Caption = Me.ListView1.ListItems(loc_key).SubItems(3)
     Me.lblUnidad.Caption = Me.ListView1.ListItems(loc_key).SubItems(4)
     
     Me.txtCantidad.Text = Me.ListView1.ListItems(loc_key).SubItems(3)

            Me.ListView1.Visible = False
            
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SP_SUBMATERIAPRIMAxPRODUCTO"
            oCmdEjec.CommandType = adCmdStoredProc
            
            Dim ORSd As ADODB.Recordset
            Me.lvSMP.ListItems.Clear
            Dim ITEMs As Object
            Set ORSd = oCmdEjec.Execute(, Array(LK_CODCIA, Me.lblCodigo.Caption))
            If Not ORSd.EOF Then
            
            Do While Not ORSd.EOF
                Set ITEMs = Me.lvSMP.ListItems.Add(, , ORSd!INSUMO)
                ITEMs.Tag = ORSd!Codigo
                ITEMs.SubItems(1) = ORSd!proporcion
                ITEMs.SubItems(2) = ORSd!UNIDAD
                ORSd.MoveNext
            Loop
            End If
            Me.txtCantidad.Enabled = True
            Me.lvSMP.Enabled = True
             Me.lvSMP.SetFocus
                End If
           
          
            
     
End Sub

Private Sub txtnro_KeyPress(KeyAscii As Integer)
If SoloNumeros(KeyAscii) Then KeyAscii = 0
If KeyAscii = vbKeyReturn Then
If Len(Trim(Me.txtNro.Text)) = 0 Then Exit Sub
If Not IsNumeric(Me.txtNro.Text) Then Exit Sub
Me.lvSMP.ListItems.Clear
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SP_CONVERSION_PRINT"
Dim orsP As ADODB.Recordset

Set orsP = oCmdEjec.Execute(, Array(Me.txtNro.Text, LK_CODCIA))

If Not orsP.EOF Then
    Me.lblMP.Caption = orsP!producto
    Me.lblUnidad.Caption = orsP!UNIDAD
    
    orsP.MoveNext
    
    Dim itemC As Object
    
    Do While Not orsP.EOF
    Set itemC = Me.lvSMP.ListItems.Add(, , orsP!producto)
    itemC.SubItems(1) = orsP!PRECIO
    itemC.SubItems(2) = orsP!UNIDAD
    itemC.SubItems(3) = orsP!Cantidad
        orsP.MoveNext
    Loop
    Me.Toolbar1.Buttons(1).Enabled = False
    Me.Toolbar1.Buttons(2).Enabled = False
   Me.Toolbar1.Buttons(3).Enabled = True
    Me.Toolbar1.Buttons(4).Enabled = True
End If
End If
End Sub
