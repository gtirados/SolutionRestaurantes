VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmProduccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes de Producción"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12450
   BeginProperty Font 
      Name            =   "Tahoma"
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
   ScaleHeight     =   6705
   ScaleWidth      =   12450
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   2280
      Width           =   7935
      Begin MSDataListLib.DataCombo DatUsuarios 
         Height          =   315
         Left            =   1920
         TabIndex        =   21
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ASIGNADO A:"
         Height          =   195
         Left            =   720
         TabIndex        =   20
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdSEARCH 
      Caption         =   "Busca OP"
      Height          =   360
      Left            =   240
      TabIndex        =   18
      Top             =   5160
      Width           =   990
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   600
      Left            =   6840
      TabIndex        =   17
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   5400
      TabIndex        =   14
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   4575
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2280
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   615
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4022
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
   Begin VB.Frame Frame1 
      Caption         =   "Dato Seleccionado"
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   7935
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO:"
         Height          =   195
         Left            =   960
         TabIndex        =   13
         Top             =   300
         Width           =   825
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCTO TERMINADO:"
         Height          =   195
         Left            =   165
         TabIndex        =   12
         Top             =   660
         Width           =   1875
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COSTO:"
         Height          =   195
         Left            =   1080
         TabIndex        =   11
         Top             =   1020
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UNIDAD:"
         Height          =   195
         Left            =   4920
         TabIndex        =   10
         Top             =   240
         Width           =   780
      End
      Begin VB.Label lblCodigo 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label lblMP 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2040
         TabIndex        =   8
         Top             =   600
         Width           =   5715
      End
      Begin VB.Label lblCosto 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2040
         TabIndex        =   7
         Top             =   960
         Width           =   1635
      End
      Begin VB.Label lblUnidad 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5880
         TabIndex        =   6
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STOCK:"
         Height          =   195
         Left            =   5040
         TabIndex        =   5
         Top             =   960
         Width           =   690
      End
      Begin VB.Label lblStock 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5880
         TabIndex        =   4
         Top             =   960
         Width           =   1875
      End
   End
   Begin MSComctlLib.ListView lvSMP 
      Height          =   2175
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   3836
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD A PROCESAR:"
      Height          =   195
      Left            =   5400
      TabIndex        =   15
      Top             =   120
      Width           =   2220
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BUSQUEDA PRODUCTO:"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1755
   End
End
Attribute VB_Name = "frmProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private loc_key  As Integer
Private Sub ConfigurarLV()
With Me.ListView1
    .ColumnHeaders.Add , , "CODIGO", 700
    .ColumnHeaders.Add , , "MATERIA PRIMA", 3000
    .ColumnHeaders.Add , , "COSTO"
    .ColumnHeaders.Add , , "STOCK"
    .ColumnHeaders.Add , , "UNIDAD"
    .FullRowSelect = True
    .Gridlines = True
    .HideColumnHeaders = True
    .View = lvwReport
    .HideSelection = False
End With

With Me.lvSMP
    
    .ColumnHeaders.Add , , "MATERIA PRIMA", 3000
    .ColumnHeaders.Add , , "PROPORCION"
    .ColumnHeaders.Add , , "UNIDAD"
    .ColumnHeaders.Add , , "STOCK"
    .ColumnHeaders.Add , , "COSTO"
    .LabelEdit = lvwManual
    .FullRowSelect = True
    .Gridlines = True
    .View = lvwReport
End With
End Sub

Private Sub cmdProcesar_Click()

    If Len(Trim(Me.lblCodigo.Caption)) = 0 Then
        MsgBox "Debe buscar el Producto.", vbCritical, Pub_Titulo
        Exit Sub
    End If

    If Me.lvSMP.ListItems.count = 0 Then
        MsgBox "No hay nada que procesar.", vbCritical, Pub_Titulo
        Exit Sub
    End If

    If Not IsNumeric(Me.txtCantidad.Text) Then
        MsgBox "La Cantidad es incorrecta.", vbInformation, Pub_Titulo
        Me.txtCantidad.SetFocus
        Me.txtCantidad.SelStart = 0
        Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
        Exit Sub
    End If

    If val(Me.txtCantidad.Text) <= 0 Then
        MsgBox "Debe ingresar la Cantidad.", vbCritical, Pub_Titulo
        Me.txtCantidad.SetFocus
        Me.txtCantidad.SelStart = 0
        Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
        Exit Sub
    End If
    
    If Me.DatUsuarios.BoundText = "" Then
        MsgBox "Debe elegir el Empleado.", vbCritical, Pub_Titulo
        Me.DatUsuarios.SetFocus
        Exit Sub
    End If

    Dim oCant As Double
    Dim f As Integer
    Dim vPasa As Boolean

    vPasa = True
    oCant = 0

    For f = 1 To Me.lvSMP.ListItems.count
        oCant = Me.lvSMP.ListItems(f).SubItems(1) * Me.txtCantidad.Text
        If oCant > Me.lvSMP.ListItems(f).SubItems(3) Then
            vPasa = False
            Exit For
        End If
    Next

' VERIFICA STOCK GTS

'    If Not vPasa Then
'        MsgBox "No pasa", vbCritical, Pub_Titulo
'        Exit Sub
'    End If

    oCant = 0

    For f = 1 To Me.lvSMP.ListItems.count
        oCant = oCant + Me.lvSMP.ListItems(f).SubItems(3)
    Next

    On Error GoTo grabar
Dim nroe As Double
Dim nros As Double
nroe = 0
nros = 0

    Pub_ConnAdo.BeginTrans
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPPROCESAR_PARTEPRODUCCION"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 20, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODART", adBigInt, adParamInput, , Me.lblCodigo.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@COSTO", adDouble, adParamInput, , oCant)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CANTIDAD", adDouble, adParamInput, , Me.txtCantidad.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@VENDEDOR", adDouble, adParamInput, , Me.DatUsuarios.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMOPERENTRADA", adDouble, adParamOutput, , nroe)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMOPERSALIDA", adDouble, adParamOutput, , nros)

    oCmdEjec.Execute
    
    nroe = oCmdEjec.Parameters("@NUMOPERSALIDA").Value

    MsgBox "Datos almacenados Correctamente.", vbInformation, Pub_Titulo
    Pub_ConnAdo.CommitTrans



If MsgBox("¿Desea imprimir el parte.?", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub

Imprime nroe, Me.lblMP.Caption, Me.txtCantidad.Text
    LimpiarControles Me
    Me.lblUnidad.Caption = ""
    Me.lblCodigo.Caption = ""
    Me.lblStock.Caption = ""
    Me.lblMP.Caption = ""
    Me.lblCosto.Caption = ""
    Me.lvSMP.ListItems.Clear
    Me.ListView1.ListItems.Clear

    Exit Sub

grabar:
    Pub_ConnAdo.RollbackTrans
    MsgBox Err.Description, vbInformation, Pub_Titulo


End Sub


Public Sub Imprime(xsalida As Double, xMateriaPrima As String, xCantidad As Double)
LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_PRODUCCION_PRINT"

'    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, IIf(Me.cboComedor.ListIndex = 1, "01", "02"))
'    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SALIDA", adDouble, adParamInput, , CsALIDA)

oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SALIDA", adDouble, adParamInput, , xsalida)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
        


    Dim ORSd As ADODB.Recordset

    Set ORSd = oCmdEjec.Execute

    Dim objCrystal  As New CRAXDRT.APPLICATION
Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions

        Dim crParamDef  As CRAXDRT.ParameterFieldDefinition
    Dim RutaReporte As String

    RutaReporte = "c:\Admin\Nordi\PRODUCCION.rpt"

    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
    
    Set crParamDefs = VReporte.ParameterFields

            For Each crParamDef In crParamDefs

                Select Case crParamDef.ParameterFieldName

                    Case "fPRODUCTO"
                        crParamDef.AddCurrentValue xMateriaPrima '
                    Case "fCANTIDAD"
                    crParamDef.AddCurrentValue CStr(xCantidad)
              
                End Select

            Next
            
            
    VReporte.DataBase.SetDataSource ORSd, 3, 1
    'frmprint.CRViewer1.ReportSource = VReporte
    'frmprint.CRViewer1.ViewReport
'    crwVisor.ReportSource = VReporte
'    crwVisor.ViewReport
    
    
    
    VReporte.PrintOut False, 1, , 1, 1
    Set objCrystal = Nothing
    Set VReporte = Nothing
End Sub

Private Sub cmdSEARCH_Click()
frmProduccion_Find.Show vbModal
End Sub

Private Sub Form_Load()
CentrarFormulario MDIForm1, Me
ConfigurarLV

LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "USP_VEMAES_LISTADO"
Dim ORSd As ADODB.Recordset
Set ORSd = oCmdEjec.Execute(, LK_CODCIA)
Set Me.DatUsuarios.RowSource = ORSd
Me.DatUsuarios.ListField = ORSd.Fields(1).Name
Me.DatUsuarios.BoundColumn = ORSd.Fields(0).Name
End Sub




Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Me.DatUsuarios.SetFocus
End If
End Sub

Private Sub txtMP_Change()
Dim ORSDatos As ADODB.Recordset

    Me.ListView1.ListItems.Clear
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPPRODUCTO_LIST"
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
        Me.txtCantidad.SetFocus
'
'        Me.txtCantidad.Text = Me.ListView1.ListItems(loc_key).SubItems(3)
'
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
                ITEMs.SubItems(3) = ORSd!stock
                ITEMs.SubItems(4) = ORSd!costo
                ORSd.MoveNext
            Loop

        End If
'
'        Me.txtCantidad.Enabled = True
'        Me.lvSMP.Enabled = True
'        Me.lvSMP.SetFocus
    End If
            
End Sub
