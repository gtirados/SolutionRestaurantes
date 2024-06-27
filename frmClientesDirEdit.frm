VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmClientesDirEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modificar Dirección"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9750
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
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   600
      Left            =   7200
      Picture         =   "frmClientesDirEdit.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   600
      Left            =   7200
      Picture         =   "frmClientesDirEdit.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtDirAdd 
      Height          =   285
      Left            =   1380
      MaxLength       =   150
      TabIndex        =   1
      Tag             =   "X"
      Top             =   105
      Width           =   8295
   End
   Begin VB.TextBox txtUrbanizacion 
      Height          =   285
      Left            =   1380
      TabIndex        =   3
      Top             =   810
      Width           =   4575
   End
   Begin VB.TextBox txtRefAdd 
      Height          =   285
      Left            =   1380
      MaxLength       =   150
      TabIndex        =   2
      Tag             =   "X"
      Top             =   450
      Width           =   8295
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1935
      Left            =   1380
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1095
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3413
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
   Begin MSDataListLib.DataCombo DatZona 
      Height          =   315
      Left            =   6780
      TabIndex        =   4
      Tag             =   "X"
      Top             =   795
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label lblIdDir 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   9120
      TabIndex        =   12
      Top             =   1440
      Width           =   555
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      Height          =   195
      Left            =   420
      TabIndex        =   11
      Top             =   150
      Width           =   870
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zona:"
      Height          =   195
      Left            =   6180
      TabIndex        =   10
      Top             =   855
      Width           =   510
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Urbanización:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   855
      Width           =   1170
   End
   Begin VB.Label lblUrb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-1"
      Height          =   255
      Left            =   9900
      TabIndex        =   8
      Top             =   825
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Referencia:"
      Height          =   195
      Left            =   300
      TabIndex        =   7
      Top             =   495
      Width           =   990
   End
End
Attribute VB_Name = "frmClientesDirEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gAcepta As Boolean
Public gIDZONA As Integer
Public gIDDIR As Integer
Public gIDURB As Integer
Public gURB As String
Public gREFERENCIA As String
Public gDIRECCION As String
Public gZONA As String
Dim loc_key  As Integer
Private vBuscar As Boolean 'variable para la busqueda de clientes

Private Sub CargarZonas()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_CLIENTE_ZONAREPARTO_LIST"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim ORSz As ADODB.Recordset

    Set ORSz = oCmdEjec.Execute(, LK_CODCIA)
    Set Me.DatZona.RowSource = ORSz
    Me.DatZona.ListField = ORSz.Fields(1).Name
    Me.DatZona.BoundColumn = ORSz.Fields(0).Name
    Me.DatZona.BoundText = -1
    
  
End Sub

Private Sub cmdAceptar_Click()
gAcepta = True
gIDDIR = Me.lblIdDir.Caption
gIDZONA = Me.DatZona.BoundText
gIDURB = Me.lblurb.Caption
gDIRECCION = Me.txtDirAdd.Text
gZONA = Me.DatZona.Text
gREFERENCIA = Me.txtRefAdd.Text

If gIDURB = -1 Then
    gURB = ""
Else
    gURB = Me.txtUrbanizacion.Text
End If

Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub DatZona_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me

End Sub

Private Sub Form_Load()
gAcepta = False
vBuscar = True
 With Me.ListView1
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .View = lvwReport
  
        .ColumnHeaders.Add , , "Cliente", 5000
  
        .MultiSelect = False
    End With
    CargarZonas
End Sub

Private Sub txtDirAdd_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtRefAdd_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtUrbanizacion_Change()
vBuscar = True
Me.lblurb.Caption = -1
End Sub

Private Sub txtUrbanizacion_KeyDown(KeyCode As Integer, Shift As Integer)
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

    If KeyCode2 = 33 Then
        loc_key = loc_key - 17

        If loc_key < 1 Then loc_key = 1
        GoTo posicion
    End If

    If KeyCode = 27 Then
        Me.ListView1.Visible = False
        Me.txtUrbanizacion.Text = ""
Me.lblurb.Caption = "-1"
    End If

    GoTo fin
posicion:
    ListView1.ListItems.Item(loc_key).Selected = True
    ListView1.ListItems.Item(loc_key).EnsureVisible
    'txtRS.Text = Trim(ListView11.ListItems.Item(loc_key).Text) & " "
    txtUrbanizacion.SelStart = Len(Me.txtUrbanizacion.Text)
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtUrbanizacion_KeyPress(KeyAscii As Integer)
  KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then
        If vBuscar Then
            Me.ListView1.ListItems.Clear
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SpListarCliUrbanizaciones"
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SEARCH", adVarChar, adParamInput, 80, Me.txtUrbanizacion.Text)
                
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
                frmClientesUrbAdd.txtUrb.Text = Me.txtUrbanizacion.Text
                frmClientesUrbAdd.Show vbModal

                If frmClientesUrbAdd.gAcepta Then
                    Me.txtUrbanizacion.Text = frmClientesUrbAdd.gNombre
                    Me.lblurb.Caption = frmClientesUrbAdd.gIde
                    Me.ListView1.Visible = False
                    vBuscar = True
                    Me.DatZona.SetFocus
                End If
            End If
        
        Else
            Me.txtUrbanizacion.Text = Me.ListView1.ListItems(loc_key).Text
            Me.lblurb.Caption = Me.ListView1.ListItems(loc_key).Tag
            
            Me.ListView1.Visible = False
            vBuscar = True
            SendKeys "{tab}"
            'Me.lvDetalle.SetFocus
        End If
    End If
End Sub
