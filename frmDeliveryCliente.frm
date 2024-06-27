VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmDeliveryCliente 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nuevo Cliente"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   9540
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
   ScaleHeight     =   3900
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtReferencia 
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   5
      Tag             =   "X"
      Top             =   1680
      Width           =   6375
   End
   Begin VB.CommandButton cmdSunat 
      Caption         =   "..."
      Height          =   360
      Left            =   8760
      TabIndex        =   11
      Top             =   480
      Width           =   630
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1455
      Left            =   7920
      TabIndex        =   10
      Top             =   2355
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   2566
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
   Begin VB.TextBox txtUrb 
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Top             =   2040
      Width           =   6375
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   720
      Left            =   7440
      Picture         =   "frmDeliveryCliente.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   720
      Left            =   6240
      Picture         =   "frmDeliveryCliente.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtRS 
      Height          =   285
      Left            =   2280
      MaxLength       =   80
      TabIndex        =   0
      Tag             =   "X"
      Top             =   165
      Width           =   6375
   End
   Begin VB.TextBox txtRuc 
      Height          =   285
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   1
      Tag             =   "X"
      Top             =   525
      Width           =   2535
   End
   Begin VB.TextBox txtDni 
      Height          =   285
      Left            =   6960
      MaxLength       =   8
      TabIndex        =   2
      Tag             =   "X"
      Top             =   525
      Width           =   1695
   End
   Begin VB.TextBox txtFono 
      Height          =   285
      Left            =   2280
      MaxLength       =   9
      TabIndex        =   3
      Tag             =   "X"
      Top             =   885
      Width           =   2535
   End
   Begin VB.TextBox txtDir 
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   4
      Tag             =   "X"
      Top             =   1245
      Width           =   6375
   End
   Begin MSDataListLib.DataCombo DatZonaR 
      Height          =   315
      Left            =   2280
      TabIndex        =   7
      Top             =   2400
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Referencia:"
      Height          =   195
      Left            =   1185
      TabIndex        =   20
      Top             =   1680
      Width           =   990
   End
   Begin VB.Label lblUrb 
      BackStyle       =   0  'Transparent
      Caption         =   "-1"
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   1680
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Urbanización:"
      Height          =   195
      Left            =   1005
      TabIndex        =   18
      Top             =   2085
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruc:"
      Height          =   195
      Left            =   1785
      TabIndex        =   17
      Top             =   570
      Width           =   390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dni:"
      Height          =   195
      Left            =   6480
      TabIndex        =   16
      Top             =   570
      Width           =   360
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono:"
      Height          =   195
      Left            =   1365
      TabIndex        =   15
      Top             =   930
      Width           =   810
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      Height          =   195
      Left            =   1305
      TabIndex        =   14
      Top             =   1290
      Width           =   870
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zona:"
      Height          =   195
      Left            =   1665
      TabIndex        =   13
      Top             =   2460
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre o Razón Social:"
      Height          =   195
      Left            =   105
      TabIndex        =   12
      Top             =   210
      Width           =   2070
   End
End
Attribute VB_Name = "frmDeliveryCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim loc_key  As Integer
Private vBuscar As Boolean 'variable para la busqueda de clientes
Public gIDz As Double
Public vTIPO As String
Public gDIRECCION As String
Public gDNI As String
Public gRS As String

Private Sub cmdAceptar_Click()

If Len(Trim(Me.txtFono.Text)) = 0 Then
    MsgBox "Debe ingresar el telefono.", vbCritical, Pub_Titulo
    Me.txtFono.SetFocus

    Exit Sub

End If

If Len(Trim(Me.txtDir.Text)) = 0 And frmDeliveryApp.chkRecojo.Value = 0 Then
    MsgBox "Debe ingresar la dirección.", vbCritical, Pub_Titulo
    Me.txtDir.SetFocus

    Exit Sub

End If

LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SP_CLIENTE_VALIDA_N"
                
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RUC", adVarChar, adParamInput, 11, Trim(Me.txtRuc.Text))
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DNI", adChar, adParamInput, 8, Trim(Me.txtDni.Text))
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FONO", adVarChar, adParamInput, 20, Me.txtFono.Text)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RS", adVarChar, adParamInput, 80, Me.txtRS.Text)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPO", adChar, adParamInput, 1, "C")
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CONTINUA", adBoolean, adParamOutput, 20, vCONTINUA)

Set orsM = oCmdEjec.Execute
                
vCONTINUA = oCmdEjec.Parameters("@CONTINUA").Value
                    
sMSN = orsM!Mensaje
                    
If Len(Trim(sMSN)) <> 0 Then
    If Not vCONTINUA Then
        MsgBox sMSN, vbCritical, Pub_Titulo

        Exit Sub

    Else

        If MsgBox(sMSN + vbCrLf + "¿Desea continuar con la operación?", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub
    End If
                    
End If
                
oCmdEjec.CommandText = "SP_CLIENTE_REGISTRAR"
    
LimpiaParametros oCmdEjec

On Error GoTo grabar
               
gIDz = -1
gIDz = 1
gDNI = ""

oCmdEjec.Prepared = True
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CLI_NOMBRE", adVarChar, adParamInput, 80, Trim(Me.txtRS.Text))
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPO", adChar, adParamInput, 1, "C")
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CLI_TELEF1", adVarChar, adParamInput, 20, Me.txtFono.Text)
                
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CLI_CASA_DIREC", adVarChar, adParamInput, 150, Me.txtDir.Text)
                
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RUC", adChar, adParamInput, 11, Me.txtRuc.Text)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DNI", adChar, adParamInput, 8, Me.txtDni.Text)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHANAC", adDBTimeStamp, adParamInput, , Date)
                
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@EMAIL", adVarChar, adParamInput, 60, "")
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@OBS", adVarChar, adParamInput, 1000, Me.txtReferencia.Text)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adInteger, adParamInput, , Me.DatZonaR.BoundText)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDURB", adBigInt, adParamInput, , Me.lblurb.Caption)
    
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamOutput, , gIDz)

oCmdEjec.Execute
'oCmdEjec.Execute , Array(CodCia, Trim(Me.txtCodigo.Text), Trim(Me.txtDenominacion.Text), Trim(Me.txtZona.Text))

gIDz = oCmdEjec.Parameters("@IDCLIENTE").Value
gDIRECCION = Me.txtDir.Text
gRS = Me.txtRS.Text
gDNI = Me.txtDni.Text
Unload Me

Exit Sub

grabar:
gIDz = -1
MsgBox Err.Description, vbInformation, Pub_Titulo
End Sub

Private Sub cmdCancelar_Click()
gIDz = -1
Unload Me
End Sub

Private Sub cmdSunat_Click()

If Len(Trim(Me.txtRuc.Text)) = 0 And Len(Trim(Me.txtDni.Text)) = 0 Then
    MsgBox "Debe ingresar el RUC o DNI.", vbCritical, Pub_Titulo
    Me.txtRuc.SetFocus
    Exit Sub
End If

If Len(Trim(Me.txtRuc.Text)) <> 0 And Len(Trim(Me.txtDni.Text)) <> 0 Then
    MsgBox "Solo debe buscar por RUC o por DNI.", vbCritical, Pub_Titulo
    Me.txtRuc.SetFocus
    Exit Sub
End If

If Len(Trim(Me.txtRuc.Text)) <> 0 Then
    If Len(Trim(Me.txtRuc.Text)) < 11 Then
        MsgBox "EL Ruc debe tener 11 dígitos.", vbCritical, Pub_Titulo
        Me.txtRuc.SetFocus
        Exit Sub
    End If
End If

If Len(Trim(Me.txtDni.Text)) <> 0 Then
    If Len(Trim(Me.txtDni.Text)) < 8 Then
        MsgBox "El DNI debe tener 8 dígitos.", vbCritical, Pub_Titulo
        Me.txtDni.SetFocus
        Exit Sub
    End If
End If


On Error GoTo cCruc

Dim p          As Object

Dim Texto      As String, xTOk As String

Dim cadena     As String, xvRUC As String

Dim sInputJson As String, xEsRuc As Boolean

xEsRuc = True

MousePointer = vbHourglass
Set httpURL = New WinHttp.WinHttpRequest
    
If Len(Trim(Me.txtDni)) <> 0 Then
       
    xEsRuc = False
        
    xvRUC = Me.txtDni.Text
Else
       
    xEsRuc = True
       
    xvRUC = Me.txtRuc.Text
End If

xTOk = Leer_Ini(App.Path & "\config.ini", "TOKEN", "")
    
If xEsRuc Then
    cadena = "http://dniruc.apisperu.com/api/v1/ruc/" & xvRUC & "?token=" & xTOk
Else
    cadena = "http://dniruc.apisperu.com/api/v1/dni/" & xvRUC & "?token=" & xTOk
 
    
End If
    
httpURL.Open "GET", cadena
httpURL.Send
    
Texto = httpURL.ResponseText

'sInputJson = "{items:" & Texto & "}"

Set p = JSON.parse(Texto)
    
If Texto = "[]" Then
    MousePointer = vbDefault
    MsgBox ("No se obtuvo resultados")
    Me.txtRuc.Text = ""
    Me.txtRS.Text = ""
    

    Exit Sub

End If

If Len(Trim(Texto)) = 0 Then
    MousePointer = vbDefault
    MsgBox ("No se obtuvo resultados")
    Me.txtRuc.Text = ""
    Me.txtRS.Text = ""
    

    Exit Sub

End If


  If xEsRuc Then
        Me.txtDir.Text = p.Item("direccion")
        Me.txtRS.Text = p.Item("razonSocial")
        Me.txtRuc.Text = p.Item("ruc")
        Me.txtDni.Text = ""
    Else
        Me.txtRuc.Text = ""
        Me.txtDir.Text = ""
        Me.txtDni.Text = p.Item("dni")
        Me.txtRS.Text = p.Item("nombres") & " " & p.Item("apellidoPaterno") & " " & p.Item("apellidoMaterno")
    End If
    
       
MousePointer = vbDefault

Exit Sub

cCruc:
MousePointer = vbDefault
MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub DatZonaR_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Me.cmdAceptar.SetFocus
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then cmdCancelar_Click
End Sub

Private Sub Form_Load()
gIDz = -1
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_CLIENTE_ZONAREPARTO_LIST"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim ORSz As ADODB.Recordset

    Set ORSz = oCmdEjec.Execute(, LK_CODCIA)
    Set Me.DatZonaR.RowSource = ORSz
    Me.DatZonaR.ListField = ORSz.Fields(1).Name
    Me.DatZonaR.BoundColumn = ORSz.Fields(0).Name
    Me.DatZonaR.BoundText = -1
    
     With Me.ListView1
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .View = lvwReport
  
        .ColumnHeaders.Add , , "Cliente", 5000
  
        .MultiSelect = False
    End With
End Sub

Private Sub txtDir_KeyPress(KeyAscii As Integer)
 KeyAscii = Mayusculas(KeyAscii)
 If KeyAscii = vbKeyReturn Then
    Me.txtReferencia.SetFocus
 End If
End Sub

Private Sub txtDni_KeyPress(KeyAscii As Integer)
If SoloNumeros(KeyAscii) Then KeyAscii = 0
 If KeyAscii = vbKeyReturn Then
    Me.txtFono.SetFocus
    Me.txtFono.SelStart = 0
    Me.txtFono.SelLength = Len(Me.txtFono.Text)
 End If
End Sub

Private Sub txtFono_KeyPress(KeyAscii As Integer)
If SoloNumeros(KeyAscii) Then KeyAscii = 0
 If KeyAscii = vbKeyReturn Then
    Me.txtDir.SetFocus
    Me.txtDir.SelStart = 0
    Me.txtDir.SelLength = Len(Me.txtDir.Text)
 End If
End Sub

Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
 KeyAscii = Mayusculas(KeyAscii)
 If KeyAscii = vbKeyReturn Then
 Me.DatZonaR.SetFocus
 End If
 
End Sub

Private Sub txtRS_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Me.txtRuc.SetFocus
        Me.txtRuc.SelStart = 0
        Me.txtRuc.SelLength = Len(Me.txtRuc.Text)
    End If
KeyAscii = Mayusculas(KeyAscii)
End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If SoloNumeros(KeyAscii) Then KeyAscii = 0

    If KeyAscii = vbKeyReturn Then
        Me.txtDni.SetFocus
        Me.txtDni.SelStart = 0
        Me.txtDni.SelLength = Len(Me.txtDni.Text)
    End If

End Sub

Private Sub txtUrb_Change()
vBuscar = True
Me.lblurb.Caption = -1
End Sub

Private Sub txtUrb_KeyDown(KeyCode As Integer, Shift As Integer)

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
        Me.txtUrb.Text = ""
Me.lblurb.Caption = "-1"
    End If

    GoTo fin
posicion:
    ListView1.ListItems.Item(loc_key).Selected = True
    ListView1.ListItems.Item(loc_key).EnsureVisible
    'txtRS.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
    txtUrb.SelStart = Len(Me.txtUrb.Text)
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtUrb_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then
        If vBuscar Then
            Me.ListView1.ListItems.Clear
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SpListarCliUrbanizaciones"
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SEARCH", adVarChar, adParamInput, 80, Me.txtUrb.Text)
                
            Dim ORSurb As ADODB.Recordset
                
            Set ORSurb = oCmdEjec.Execute

            Dim Item As Object
        
            If Not ORSurb.EOF Then

                Do While Not ORSurb.EOF
                    Set Item = Me.ListView1.ListItems.Add(, , ORSurb!nom)
                    Item.Tag = ORSurb!IDE
                    ORSurb.MoveNext
                Loop

                Me.ListView1.Visible = True
                Me.ListView1.ListItems(1).Selected = True
                loc_key = 1
                Me.ListView1.ListItems(1).EnsureVisible
                vBuscar = False
            Else

                If MsgBox("Urbanización no encontrada." & vbCrLf & "¿Desea agregarla?", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub
                frmClientesUrbAdd.txtUrb.Text = Me.txtUrb.Text
                frmClientesUrbAdd.Show vbModal

                If frmClientesUrbAdd.gAcepta Then
                    Me.txtUrb.Text = frmClientesUrbAdd.gNombre
                    Me.lblurb.Caption = frmClientesUrbAdd.gIde
                    Me.ListView1.Visible = False
                    vBuscar = True
                    Me.DatZonaR.SetFocus
                End If
            End If
        
        Else
            Me.txtUrb.Text = Me.ListView1.ListItems(loc_key).Text
            Me.lblurb.Caption = Me.ListView1.ListItems(loc_key).Tag
            
            Me.ListView1.Visible = False
            vBuscar = True
            Me.DatZonaR.SetFocus
            'Me.lvDetalle.SetFocus
        End If
    End If
End Sub
