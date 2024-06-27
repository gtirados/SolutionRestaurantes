VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmVentas2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Punto de Venta Carniceria"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   8775
   Begin VB.CommandButton cmdSunat 
      Caption         =   "Sunat"
      Height          =   360
      Left            =   7320
      TabIndex        =   24
      Top             =   960
      Width           =   990
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1815
      Left            =   1320
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3201
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
   Begin MSComctlLib.ListView lvVenta 
      Height          =   3135
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   7815
      _ExtentX        =   13785
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
   Begin VB.TextBox txtProducto 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   2250
      Width           =   6615
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   7815
      Begin VB.TextBox txtNro 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   9
         Top             =   270
         Width           =   1095
      End
      Begin VB.CheckBox chkEdit 
         Caption         =   "Edit"
         Height          =   255
         Left            =   4440
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin MSDataListLib.DataCombo DatTiposDoctos 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   255
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   "TiposDoctos"
      End
      Begin VB.Label lblSerie 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2280
         TabIndex        =   22
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.TextBox txtDireccion 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      TabIndex        =   6
      Top             =   1680
      Width           =   6015
   End
   Begin VB.TextBox txtRuc 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      TabIndex        =   5
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Del"
      Height          =   360
      Left            =   8040
      TabIndex        =   4
      Top             =   3240
      Width           =   630
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "&Limpiar"
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
      Left            =   6480
      TabIndex        =   3
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
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
      Left            =   5040
      TabIndex        =   2
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox txtCliente 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   6015
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCTO:"
      Height          =   195
      Left            =   225
      TabIndex        =   23
      Top             =   2295
      Width           =   1065
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECCIÓN:"
      Height          =   195
      Left            =   180
      TabIndex        =   20
      Top             =   1770
      Width           =   1110
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC:"
      Height          =   195
      Left            =   840
      TabIndex        =   19
      Top             =   1410
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Igv:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3240
      TabIndex        =   18
      Top             =   5880
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5520
      TabIndex        =   17
      Top             =   5880
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   16
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblIgv 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3840
      TabIndex        =   15
      Top             =   5835
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6360
      TabIndex        =   14
      Top             =   5835
      Width           =   1590
   End
   Begin VB.Label lblSubTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   5835
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTE:"
      Height          =   195
      Left            =   480
      TabIndex        =   12
      Top             =   1035
      Width           =   810
   End
End
Attribute VB_Name = "frmVentas2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim loc_key  As Integer
Private vPagActEnt, vPagTotEnt As Integer
Private vIniLeft As Integer
Private vIniTop As Integer
Private vBuscar As Boolean 'variable para la busqueda de clientes
Private ORStd As ADODB.Recordset 'VARIABLE PARA SABER SI EL TIPO DE DOCUMENTO ES EDITABLE

Private Sub chkEdit_Click()
Me.txtNro.Enabled = Me.chkEdit.Value
If Me.chkEdit.Value Then
Me.txtNro.SetFocus
Me.txtNro.SelStart = 0
Me.txtNro.SelLength = Len(Me.txtNro.Text)
End If
End Sub

Private Sub cmdDel_Click()
If Me.lvVenta.ListItems.count = 0 Then Exit Sub
Me.lvVenta.ListItems.Remove Me.lvVenta.SelectedItem.Index
sumatoria
End Sub

Private Sub cmdGrabar_Click()

    If Me.DatTiposDoctos.BoundText = "01" And Len(Trim(Me.txtRuc.Text)) = 0 Then
        MsgBox "Debe ingresar el cliente para una factura.", vbCritical, Pub_Titulo

        Exit Sub

    End If

    If Me.lvVenta.ListItems.count = 0 Then
        MsgBox "Debe agregar productos.", vbInformation, Pub_Titulo

        Exit Sub

    End If

    On Error GoTo xGraba

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_ENTRADAS_REGISTRAR"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim Xitem As Object

    Dim xDET  As String

    xDET = ""

    If Me.lvVenta.ListItems.count <> 0 Then
        xDET = "<r>"

        For Each Xitem In Me.lvVenta.ListItems

            xDET = xDET + "<d "
            xDET = xDET + "cp=""" & Xitem.Tag & """ "
            xDET = xDET + "st=""" & Xitem.Text & """ "
            xDET = xDET + "pr=""" & Xitem.SubItems(2) & """ "
            xDET = xDET + "un=""" & "UND" & """ "
            xDET = xDET + "sc=""" & Xitem.Index & """ "
            xDET = xDET + "cam=""" & "" & """ "
            xDET = xDET + "pz=""" & Xitem.SubItems(4) & """ "
            xDET = xDET + "/>"
        Next

        xDET = xDET + "</r>"
    End If

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SerDoc", adVarChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NroDoc", adInteger, adParamInput, , Me.txtNro.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FBG", adChar, adParamInput, 1, Left(Me.DatTiposDoctos.Text, 1))

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@XMLDET", adVarChar, adParamInput, 4000, xDET)

    If Me.DatTiposDoctos.BoundText = "01" Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcli", adInteger, adParamInput, , Me.txtRuc.Tag)
    Else
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcli", adInteger, adParamInput, , 1)
    End If

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@totalfac", adDouble, adParamInput, , Me.lblTotal.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@diascre", adDouble, adParamInput, , 0)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@farjabas", adInteger, adParamInput, , 0)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODIGODOCTO", adChar, adParamInput, 2, Me.DatTiposDoctos.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@AUTONUMFAC", adInteger, adParamOutput, , 0)    'gts agregado
    oCmdEjec.Execute
    
    Me.txtNro.Text = oCmdEjec.Parameters("@AUTONUMFAC").Value   'gts agregado
 
    If MsgBox("¿Desea imprimir?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
        Imprimir Left(Me.DatTiposDoctos.Text, 1), False
    End If
    
    Dim Xpzas As Boolean

    Xpzas = False

    For Each Xitem In Me.lvVenta.ListItems

        If Xitem.SubItems(4) <> 0 Then
            Xpzas = True
            Exit For
        End If

    Next
    
    If Xpzas Then
        Imprimir2 Me.lblSerie.Caption, Me.txtNro.Text, LK_FECHA_DIA, Left(Me.DatTiposDoctos.Text, 1)
    End If
    
    Me.lvVenta.ListItems.Clear
    Me.txtCliente.Text = ""
    Me.txtDireccion.Text = ""
    Me.lblSubTotal.Caption = "0.00"
    Me.lblIgv.Caption = "0.00"
    Me.txtRuc.Text = ""
    Me.lblTotal.Caption = "0.00"
    
      LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_TIPOS_DOCTOS_LIST"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    Set ORStd = oCmdEjec.Execute
    Set Me.DatTiposDoctos.RowSource = ORStd
    Me.DatTiposDoctos.ListField = ORStd.Fields(1).Name
    Me.DatTiposDoctos.BoundColumn = ORStd.Fields(0).Name
    
    Me.txtDireccion.Text = ""
    
     If ORStd.RecordCount <> 0 Then
        Me.DatTiposDoctos.BoundText = ORStd!Codigo
    
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SP_SERIES_CARGAR"
        oCmdEjec.CommandType = adCmdStoredProc

        Dim xSerie As String

        Dim xNro   As Double

        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, LK_CODUSU)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FBG", adChar, adParamInput, 1, Left(Me.DatTiposDoctos.Text, 1)) ' Me.cboTipoDocto.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIE", adChar, adParamOutput, 3, 1)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MAXIMO", adBigInt, adParamOutput, , 1)
        oCmdEjec.Execute

        xSerie = oCmdEjec.Parameters("@SERIE").Value
        'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)

        Me.lblSerie.Caption = xSerie
        Me.txtNro.Text = oCmdEjec.Parameters("@MAXIMO").Value
        ORStd.Filter = "CODIGO='" & Me.DatTiposDoctos.BoundText & "'"

        If ORStd.RecordCount <> 0 Then
            Me.chkEdit.Enabled = ORStd!Editable
        End If

        ORStd.Filter = ""
    End If
    
    Exit Sub

xGraba:
    MsgBox Err.Description
End Sub

Private Sub cmdLimpiar_Click()
Limpiar
End Sub



Private Sub Command1_Click()

End Sub

Private Sub cmdSunat_Click()

    If Len(Trim(Me.txtRuc.Text)) = 0 Then
        If IsNumeric(Me.txtCliente.Text) Then
            If Len(Trim(Me.txtCliente.Text)) = 11 Then
                frmFacComandaSunat.gRUC = Trim(Me.txtCliente.Text)
                frmFacComandaSunat.Show vbModal

                If frmFacComandaSunat.gAcepta Then
                    Me.txtDireccion.Text = Trim(frmFacComandaSunat.gDIR)
                    Me.txtCliente.Text = Trim(frmFacComandaSunat.gRS)
                    Me.txtRuc.Text = frmFacComandaSunat.gRUC
                    LimpiaParametros oCmdEjec
                    oCmdEjec.CommandText = "SP_CLIENTES_UPDATE_DATOS_SUNAT"
                    oCmdEjec.CommandType = adCmdStoredProc
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
            
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RAZONSOCIAL", adVarChar, adParamInput, 60, frmFacComandaSunat.gRS)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DIRECCION", adVarChar, adParamInput, 50, Left(frmFacComandaSunat.gDIR, 50))
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RUC", adVarChar, adParamInput, 11, frmFacComandaSunat.gRUC)
                   oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamOutput, , 0)
                    oCmdEjec.Execute
                    If Not IsNull(oCmdEjec.Parameters("@IDCLIENTE").Value) Then
                        Me.txtRuc.Tag = oCmdEjec.Parameters("@IDCLIENTE").Value
                    End If
                End If

            Else
                MsgBox "El ruc debe tener 11 caracteres", vbCritical, Pub_Titulo
            End If

        Else
            MsgBox "El ruc debe ser Numeros", vbCritical, Pub_Titulo
        End If

    Else
        frmFacComandaSunat.gRUC = Trim(Me.txtRuc.Text)
        frmFacComandaSunat.Show vbModal

        If frmFacComandaSunat.gAcepta Then
            Me.txtDireccion.Text = Trim(frmFacComandaSunat.gDIR)
            Me.txtCliente.Text = Trim(frmFacComandaSunat.gRS)
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SP_CLIENTES_UPDATE_DATOS_SUNAT"
            oCmdEjec.CommandType = adCmdStoredProc
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
            
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RAZONSOCIAL", adVarChar, adParamInput, 60, frmFacComandaSunat.gRS)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DIRECCION", adVarChar, adParamInput, 50, Left(frmFacComandaSunat.gDIR, 50))
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RUC", adVarChar, adParamInput, 11, frmFacComandaSunat.gRUC)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, 10, Me.txtRuc.Tag)
            oCmdEjec.Execute
            'MsgBox Me.txtRuc.Tag
        End If
    End If
End Sub

Private Sub DatTiposDoctos_Click(Area As Integer)
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_SERIES_CARGAR"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim xSerie As String

    Dim xNro   As Double

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FBG", adChar, adParamInput, 1, Left(Me.DatTiposDoctos.Text, 1)) ' Me.cboTipoDocto.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIE", adChar, adParamOutput, 3, 1)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MAXIMO", adBigInt, adParamOutput, , 1)
    oCmdEjec.Execute

    xSerie = oCmdEjec.Parameters("@SERIE").Value
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)

    Me.lblSerie.Caption = xSerie
    Me.txtNro.Text = oCmdEjec.Parameters("@MAXIMO").Value

    ORStd.Filter = "CODIGO='" & Me.DatTiposDoctos.BoundText & "'"

    If ORStd.RecordCount <> 0 Then
        Me.chkEdit.Enabled = ORStd!Editable
    End If

    ORStd.Filter = ""
    Me.txtNro.Enabled = False
    Me.chkEdit.Value = False
End Sub

Private Sub Form_Load()
    ConfiguraLV
    CenterMe Me
   vBuscar = True
    'CargarNumeracion "B"
    
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_TIPOS_DOCTOS_LIST"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    Set ORStd = oCmdEjec.Execute
    Set Me.DatTiposDoctos.RowSource = ORStd
    Me.DatTiposDoctos.ListField = ORStd.Fields(1).Name
    Me.DatTiposDoctos.BoundColumn = ORStd.Fields(0).Name

    If ORStd.RecordCount <> 0 Then
        Me.DatTiposDoctos.BoundText = ORStd!Codigo
    
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SP_SERIES_CARGAR"
        oCmdEjec.CommandType = adCmdStoredProc

        Dim xSerie As String

        Dim xNro   As Double

        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, LK_CODUSU)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FBG", adChar, adParamInput, 1, Left(Me.DatTiposDoctos.Text, 1)) ' Me.cboTipoDocto.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIE", adChar, adParamOutput, 3, 1)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MAXIMO", adBigInt, adParamOutput, , 1)
        oCmdEjec.Execute

        xSerie = oCmdEjec.Parameters("@SERIE").Value
        'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)

        Me.lblSerie.Caption = xSerie
        Me.txtNro.Text = oCmdEjec.Parameters("@MAXIMO").Value
        ORStd.Filter = "CODIGO='" & Me.DatTiposDoctos.BoundText & "'"

        If ORStd.RecordCount <> 0 Then
            Me.chkEdit.Enabled = ORStd!Editable
        End If

        ORStd.Filter = ""
    End If
    
End Sub

Private Sub ConfiguraLV()
With Me.lvVenta

    .Gridlines = True
    .LabelEdit = lvwManual
    .FullRowSelect = True
    .View = lvwReport
    
    .HideSelection = False
    .ColumnHeaders.Add , , "Cant."
    .ColumnHeaders.Add , , "Entrada"
    .ColumnHeaders.Add , , "Precio"
    .ColumnHeaders.Add , , "Importe"
    .ColumnHeaders.Add , , "Pieza", 0


End With

With Me.ListView1
    .FullRowSelect = True
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .ColumnHeaders.Add , , "Codigo", 1000
    .ColumnHeaders.Add , , "Cliente", 5000
    .ColumnHeaders.Add , , "Ruc", 0
    .ColumnHeaders.Add , , "Direcion", 0
    .MultiSelect = False
End With
End Sub

Private Sub sumatoria()

    Dim itemX As Object

    Dim vT    As Double

    vT = 0

    For Each itemX In Me.lvVenta.ListItems

        vT = vT + itemX.SubItems(3)
    Next
 
    Me.lblTotal.Caption = FormatCurrency(vT)
    Me.lblSubTotal.Caption = FormatCurrency(Round(CDec(vT) / CDec((LK_IGV / 100) + 1), 2))
    Me.lblIgv.Caption = FormatCurrency(CDec(vT) - CDec(Me.lblSubTotal.Caption), 2)

End Sub

Private Sub Limpiar()
Me.lblIgv.Caption = FormatCurrency(0)
Me.lblSubTotal.Caption = FormatCurrency(0)
Me.lblTotal.Caption = FormatCurrency(0)
Me.lvVenta.ListItems.Clear
Me.txtCliente.Text = ""

End Sub


Private Sub lvVenta_DblClick()
frmVentas2piezas.Show vbModal
If frmVentas2piezas.gAcepta Then
    Me.lvVenta.SelectedItem.SubItems(4) = frmVentas2piezas.gPieza
End If
End Sub

'Private Sub CargarNumeracion(dTipo As String)
'LimpiaParametros oCmdEjec
'oCmdEjec.CommandText = "SP_SERIES_CARGAR"
'oCmdEjec.CommandType = adCmdStoredProc
'
'Dim XsERIE As String
'Dim Xnro As Double
'
'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, LK_CODUSU)
'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FBG", adChar, adParamInput, 1, dTipo)
'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIE", adInteger, adParamOutput, 200, 1)
'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MAXIMO", adBigInt, adParamOutput, , 1)
'oCmdEjec.Execute
'
'XsERIE = oCmdEjec.Parameters("@SERIE").Value
''oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)
'
'Me.lblSerie.Caption = XsERIE
'Me.txtNumero.Text = oCmdEjec.Parameters("@MAXIMO").Value
'End Sub

'Private Sub optP_Click()
'Me.lblIgv.Visible = False
'Me.lblSubTotal.Visible = False
'Me.Label3.Visible = False
'Me.Label5.Visible = False
'CargarNumeracion "P"
'End Sub

Private Sub txtCliente_Change()
vBuscar = True
Me.txtRuc.Text = ""
Me.txtDireccion.Text = ""
End Sub

Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo SALE

    Dim strFindMe As String

    Dim itmFound  As Object ' ListItem    ' Variable FoundItem.

    If KeyCode = 40 Then  ' flecha abajo
        loc_key = loc_key + 1

        If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
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

        If loc_key < 1 Then loc_key = 1
        GoTo posicion
    End If

    If KeyCode = 27 Then
        Me.ListView1.Visible = False
        Me.txtCliente.Text = ""
        Me.txtRuc.Text = ""
        Me.txtDireccion.Text = ""
    End If

    GoTo fin
posicion:
    ListView1.ListItems.Item(loc_key).Selected = True
    ListView1.ListItems.Item(loc_key).EnsureVisible
    'txtRS.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
    txtCliente.SelStart = Len(Me.txtCliente.Text)
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
   KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then
        If vBuscar Then
            Me.ListView1.ListItems.Clear
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SpListarCliProv"
            Set oRsPago = oCmdEjec.Execute(, Array(LK_CODCIA, "C", Me.txtCliente.Text))

            Dim Item As Object
        
            If Not oRsPago.EOF Then

                Do While Not oRsPago.EOF
                    Set Item = Me.ListView1.ListItems.Add(, , oRsPago!CodClie)
                    Item.SubItems(1) = Trim(oRsPago!Nombre)
                    Item.SubItems(2) = IIf(IsNull(oRsPago!RUC), "", oRsPago!RUC)
                    Item.SubItems(3) = Trim(oRsPago!dir)
                    oRsPago.MoveNext
                Loop

                Me.ListView1.Visible = True
                Me.ListView1.ListItems(1).Selected = True
                loc_key = 1
                Me.ListView1.ListItems(1).EnsureVisible
                vBuscar = False
            Else

                If MsgBox("Cliente no existe." + vbCrLf + "¿Desea Crearlo.?", vbQuestion + vbYesNo, "Restaurantes") = vbYes Then
                    frmCLI.Show vbModal
                End If
                Me.ListView1.Visible = False
            End If
        
        Else
         Me.txtCliente.Text = Me.ListView1.ListItems(loc_key).SubItems(1)
            Me.txtRuc.Tag = Me.ListView1.ListItems(loc_key)
            Me.txtRuc.Text = Me.ListView1.ListItems(loc_key).SubItems(2)
            Me.txtDireccion.Text = Me.ListView1.ListItems(loc_key).SubItems(3)
            Me.ListView1.Visible = False
            Me.lvVenta.SetFocus
        End If
    End If
End Sub

Private Sub Imprimir(TipoDoc As String, Esconsumo As Boolean)
 Dim xCONTINUA      As Boolean
Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions

    Dim crParamDef  As CRAXDRT.ParameterFieldDefinition

    Dim objCrystal  As New CRAXDRT.Application

    Dim vIgv        As Currency

    Dim vSubTotal   As Currency

    Dim RutaReporte As String

    Dim xARCENCONTRADO As Boolean

    xCONTINUA = False
    xARCENCONTRADO = False
    
         LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SP_ARCHIVO_PRINT"
            oCmdEjec.CommandType = adCmdStoredProc
    
            Dim ORSd        As ADODB.Recordset

            

            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODIGO", adChar, adParamInput, 2, Me.DatTiposDoctos.BoundText)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@COMSUMO", adBoolean, adParamInput, , False)
    
            Set ORSd = oCmdEjec.Execute
            RutaReporte = PUB_RUTA_REPORTE & ORSd!ReportE
    
            If ORSd!ReportE = "" Then
                If MsgBox("El Archivo no existe." & vbCrLf & "¿Desea continuar sin imprimir?.", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
                    xCONTINUA = True
                End If

            Else
                FileName = dir(RutaReporte)

                If FileName = "" Then
                    If MsgBox("El Archivo no existe." & vbCrLf & "¿Desea continuar sin imprimir?.", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
                        xCONTINUA = True
                    Else

                        Exit Sub

                    End If

                Else
                    xCONTINUA = True
                    xARCENCONTRADO = True
                    'ImprimirDocumentoVenta Me.DatTiposDoctos.BoundText, Me.DatTiposDoctos.Text, Me.chkConsumo.Value, "009", 2, 100, "33", "43", "343"
                End If
            End If

            If xCONTINUA Then
              On Error GoTo printe

    
'    If Esconsumo Then
'        If TipoDoc = "B" Then
'            RutaReporte = "C:\Admin\Nordi\BolCon.rpt"
'        ElseIf TipoDoc = "F" Then
'            RutaReporte = "C:\Admin\Nordi\FacCon.rpt"
'            vSubTotal = Round((Me.lblTotal.Caption / ((100 + LK_IGV) / 100)), 2)
'            vIgv = Me.lblTotal.Caption - vSubTotal
'        End If
'
'    Else
'
'        If TipoDoc = "B" Then
'            RutaReporte = "C:\Admin\Nordi\BolDet.rpt"
'        ElseIf TipoDoc = "F" Then
'            RutaReporte = "C:\Admin\Nordi\FacDet.rpt"
'            vSubTotal = Round((Me.lblTotal.Caption / ((100 + LK_IGV) / 100)), 2)
'            vIgv = Me.lblTotal.Caption - vSubTotal
'        End If
'    End If

    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
    Set crParamDefs = VReporte.ParameterFields

    For Each crParamDef In crParamDefs

        Select Case crParamDef.ParameterFieldName

            Case "cliente"
                crParamDef.AddCurrentValue IIf(Len(Trim(Me.txtCliente.Text)) = 0, Trim(Me.txtCliente.Text), Trim(Me.txtCliente.Text))
            Case "FechaEmi"
                crParamDef.AddCurrentValue LK_FECHA_DIA
            Case "Son"
                crParamDef.AddCurrentValue CONVER_LETRAS(Me.lblTotal.Caption, "S")
            Case "total"
                crParamDef.AddCurrentValue FormatNumber(Me.lblTotal.Caption, 2)
            Case "subtotal"
                crParamDef.AddCurrentValue CStr(vSubTotal)
            Case "igv"
                crParamDef.AddCurrentValue CStr(vIgv)
            Case "SerFac"
                crParamDef.AddCurrentValue Me.lblSerie.Caption
            Case "NumFac"
                crParamDef.AddCurrentValue Me.txtNro.Text
            Case "DirClie"
                crParamDef.AddCurrentValue Me.txtDireccion.Text
            Case "RucClie"
                crParamDef.AddCurrentValue Me.txtRuc.Text
            Case "Importe" 'linea nueva
                crParamDef.AddCurrentValue FormatNumber(Me.lblTotal.Caption, 2)  'linea nueva
        End Select

    Next

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandType = adCmdStoredProc

    'oCmdEjec.CommandText = "SpPrintComanda"

    Dim rsd As ADODB.Recordset

    oCmdEjec.CommandText = "SpPrintFacturacion"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Serie", adChar, adParamInput, 3, Me.lblSerie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@nro", adInteger, adParamInput, , Me.txtNro.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fbg", adChar, adParamInput, 1, Left(Me.DatTiposDoctos.Text, 1))
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NroCom", adInteger, adParamInput, , vNroCom)

    Set rsd = oCmdEjec.Execute

    'COCINA
    'rsd.Filter = "PED_FAMILIA=2"
    Dim DD As ADODB.Recordset

    ' For i = 0 To Printers.count - 1
    '        MsgBox Printers(i).DeviceName
    '    Next
    If Not rsd.EOF Then

        VReporte.DataBase.SetDataSource rsd, 3, 1 'lleno el objeto reporte
        'VReporte.SelectPrinter Printer.DriverName, "\\laptop\doPDF v6", Printer.Port
        '
        VReporte.PrintOut False, 1, , 1, 1
        frmVisor.cr.ReportSource = VReporte
        'frmVisor.cr.ViewReport
        'frmVisor.Show vbModal
    
    End If



    Set objCrystal = Nothing
    Set VReporte = Nothing

'   LimpiaParametros oCmdEjec
'    oCmdEjec.CommandText = "SP_TIPOS_DOCTOS_LIST"
'    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
'    Set ORStd = oCmdEjec.Execute
'    Set Me.DatTiposDoctos.RowSource = ORStd
'    Me.DatTiposDoctos.ListField = ORStd.Fields(1).Name
'    Me.DatTiposDoctos.BoundColumn = ORStd.Fields(0).Name
    Exit Sub

printe:
    MostrarErrores Err
            End If

   
            

  

End Sub


Private Sub Imprimir2(xSerie As String, xNro As Double, XfECHA As Date, xFBG As String)

    On Error GoTo printe

    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions

    Dim crParamDef  As CRAXDRT.ParameterFieldDefinition

    Dim objCrystal  As New CRAXDRT.Application

    Dim vIgv        As Currency

    Dim vSubTotal   As Currency

    Dim RutaReporte As String

   
            RutaReporte = "C:\Admin\Nordi\PIEZAS.rpt"
     

    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
'    Set crParamDefs = VReporte.ParameterFields
'
'    For Each crParamDef In crParamDefs
'
'        Select Case crParamDef.ParameterFieldName
'
'            Case "cliente"
'                crParamDef.AddCurrentValue IIf(Len(Trim(Me.txtCliente.Text)) = 0, Trim(Me.txtCliente.Text), Trim(Me.txtCliente.Text))
'            Case "FechaEmi"
'                crParamDef.AddCurrentValue LK_FECHA_DIA
'            Case "Son"
'                crParamDef.AddCurrentValue CONVER_LETRAS(Me.lblTotal.Caption, "S")
'            Case "total"
'                crParamDef.AddCurrentValue FormatNumber(Me.lblTotal.Caption, 2)
'            Case "subtotal"
'                crParamDef.AddCurrentValue CStr(vSubTotal)
'            Case "igv"
'                crParamDef.AddCurrentValue CStr(vIgv)
'            Case "SerFac"
'                crParamDef.AddCurrentValue Me.lblSerie.Caption
'            Case "NumFac"
'                crParamDef.AddCurrentValue Me.txtNro.Text
'            Case "DirClie"
'                crParamDef.AddCurrentValue Me.txtDireccion.Text
'            Case "RucClie"
'                crParamDef.AddCurrentValue Me.txtRuc.Text
'            Case "Importe" 'linea nueva
'                crParamDef.AddCurrentValue FormatNumber(Me.lblTotal.Caption, 2)  'linea nueva
'        End Select
'
'    Next

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandType = adCmdStoredProc

    'oCmdEjec.CommandText = "SpPrintComanda"

    Dim rsd As ADODB.Recordset

    oCmdEjec.CommandText = "SP_VENTAS_PRINT"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adChar, adParamInput, 3, xSerie)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adDouble, adParamInput, , xNro)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , XfECHA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fbg", adChar, adParamInput, 1, xFBG)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NroCom", adInteger, adParamInput, , vNroCom)

    Set rsd = oCmdEjec.Execute

    'COCINA
    'rsd.Filter = "PED_FAMILIA=2"
    Dim DD As ADODB.Recordset

    ' For i = 0 To Printers.count - 1
    '        MsgBox Printers(i).DeviceName
    '    Next
    If Not rsd.EOF Then

        VReporte.DataBase.SetDataSource rsd, 3, 1 'lleno el objeto reporte
        'VReporte.SelectPrinter Printer.DriverName, "\\laptop\doPDF v6", Printer.Port
        '
       ' VReporte.PrintOut False, 1, , 1, 1
        frmVisor.cr.ReportSource = VReporte
        frmVisor.cr.ViewReport
        frmVisor.Show vbModal
    
    End If



    Set objCrystal = Nothing
    Set VReporte = Nothing

    Exit Sub

printe:
    MostrarErrores Err


End Sub

Private Sub txtProducto_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Len(Trim(Me.txtProducto.Text)) = 0 Then Exit Sub
        If Not IsNumeric(Me.txtProducto.Text) Then Exit Sub

        'desglozar
        Dim Codigo As String

        Dim Peso   As String

        Codigo = Left(Me.txtProducto.Text, 7)
        
        Peso = Mid(Me.txtProducto.Text, 8, 5)
        'MsgBox Codigo
        'MsgBox Peso
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SP_PRODUCTO_OBTENER_DATOS"
        oCmdEjec.CommandType = adCmdStoredProc
        
        Dim orsP As ADODB.Recordset
        
        Set orsP = oCmdEjec.Execute(, Array(LK_CODCIA, Codigo))
        
        If Not orsP.EOF Then
            If orsP.RecordCount = 0 Then
                MsgBox "No se encuentra el producto.", vbInformation, Pub_Titulo
                Exit Sub
            End If

           ' frmVentas2piezas.Show vbModal

           ' If frmVentas2piezas.gAcepta Then
  Dim itemX As Object

                Set itemX = Me.lvVenta.ListItems.Add(, , Left(Peso, 2) + "." + Right(Peso, 3))
                itemX.Tag = orsP!IDE
                itemX.SubItems(1) = orsP!producto
                itemX.SubItems(2) = orsP!PRECIO
                itemX.SubItems(3) = Round(orsP!PRECIO * CDec(Left(Peso, 2) + "." + Right(Peso, 3)), 2)
                itemX.SubItems(4) = frmVentas2piezas.gPieza
          '  End If

          
        End If

        sumatoria
        Me.txtProducto.Text = ""
    End If

End Sub
