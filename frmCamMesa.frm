VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmCamMesa 
   BackColor       =   &H8000000C&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleccione otra Mesa"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6345
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
   ScaleHeight     =   4770
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvMesas 
      Height          =   4455
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   7858
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "C&errar"
      Height          =   735
      Left            =   5280
      Picture         =   "frmCamMesa.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "&Posterior"
      Height          =   735
      Left            =   5280
      Picture         =   "frmCamMesa.frx":07AA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdAnt 
      Caption         =   "&Anterior"
      Height          =   735
      Left            =   5280
      Picture         =   "frmCamMesa.frx":0E94
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "A&ceptar"
      Height          =   735
      Left            =   5280
      Picture         =   "frmCamMesa.frx":157E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
End
Attribute VB_Name = "frmCamMesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vSeleccion As Boolean
Public vCodzon As Integer
Public tCodMesa As String
Public tMesa As String

Private Sub ConfigurarLV()
With Me.lvMesas
    .FullRowSelect = True
    .Gridlines = True
    .HideSelection = False
    .LabelEdit = lvwManual
    .View = lvwReport
    .ColumnHeaders.Add , , "Mesa", 2000
    .ColumnHeaders.Add , , "Zona", 1000
    .ColumnHeaders.Add , , "Estado", 800
End With
End Sub

Private Sub cmdAceptar_Click()
    vSeleccion = True
    tCodMesa = Me.lvMesas.SelectedItem.Tag
    tMesa = Me.lvMesas.SelectedItem.Text
    Unload Me
End Sub

Private Sub Desplzarse(vUP As Boolean)
If vUP Then
    If Me.lvMesas.SelectedItem.Index = 1 Then Exit Sub
    Me.lvMesas.ListItems(Me.lvMesas.SelectedItem.Index - 1).Selected = True
    'Me.lvMozo.SelectedItem.Index = Me.lvMozo.SelectedItem.Index - 1
Else
    If Me.lvMesas.SelectedItem.Index = Me.lvMesas.ListItems.count Then Exit Sub
    Me.lvMesas.ListItems(Me.lvMesas.SelectedItem.Index + 1).Selected = True
End If
End Sub
Private Sub cmdAnt_Click()
Desplzarse True
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdPost_Click()
Desplzarse False
End Sub

Private Sub Form_Load()
    InhabilitarCerrar Me
    vSeleccion = False
    ConfigurarLV
    CargarMesasLibresxZona
End Sub

Private Sub CargarMesasLibresxZona()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_MESAS_LISTbyCHANGE"

    With oCmdEjec
        .Parameters.Append .CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
        '.Parameters.Append .CreateParameter("@Codzon", adInteger, adParamInput, , vCodzon)
        '.Parameters.Append .CreateParameter("@Estado", adChar, adParamInput, 1, "L")
    End With

    Dim orsM As ADODB.Recordset

    Set orsM = oCmdEjec.Execute

    Do While Not orsM.EOF

        With Me.lvMesas.ListItems.Add(, , Trim(orsM!mesa))
            .Tag = Trim(orsM!codmesa)
            .SubItems(1) = Trim(orsM!ZONA)
            .SubItems(2) = Trim(orsM!ESTADO)
        End With

        'Set ItemM = Me.lvMesas.ListItems.Add(, , Trim(oRsM!Mesa))
        orsM.MoveNext
    Loop

End Sub
