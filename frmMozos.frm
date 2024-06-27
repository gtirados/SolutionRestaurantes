VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMozos 
   BackColor       =   &H8000000C&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asignar Mozo"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6030
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
   ScaleHeight     =   5025
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   975
      Left            =   4680
      Picture         =   "frmMozos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   975
      Left            =   4680
      Picture         =   "frmMozos.frx":07AA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdPosterior 
      Caption         =   "&Posterior"
      Height          =   975
      Left            =   4680
      Picture         =   "frmMozos.frx":0E95
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "A&nterior"
      Height          =   975
      Left            =   4680
      Picture         =   "frmMozos.frx":157F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvMozo 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   8493
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmMozos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vCodMoz As Integer
Public vMozo As String
Public vAcepta As Boolean
Private Sub Desplazarse(vUP As Boolean)
If Me.lvMozo.ListItems.count > 0 Then
If vUP Then
    If Me.lvMozo.SelectedItem.Index = 1 Then Exit Sub
    Me.lvMozo.ListItems(Me.lvMozo.SelectedItem.Index - 1).Selected = True
    'Me.lvMozo.SelectedItem.Index = Me.lvMozo.SelectedItem.Index - 1
Else
    If Me.lvMozo.SelectedItem.Index = Me.lvMozo.ListItems.count Then Exit Sub
    Me.lvMozo.ListItems(Me.lvMozo.SelectedItem.Index + 1).Selected = True
End If
End If
End Sub

Private Sub EnviaDatos()
    vAcepta = True
    vCodMoz = Me.lvMozo.SelectedItem.Tag
    vMozo = Me.lvMozo.SelectedItem.Text
    Unload Me
End Sub

Private Sub CargarMozos()

    Dim oRsMozos As ADODB.Recordset

    Dim itemM    As ListItem

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpListarMozos"
    Set oRsMozos = oCmdEjec.Execute(, LK_CODCIA)

    Do While Not oRsMozos.EOF

        With Me.lvMozo.ListItems.Add(, , Trim(oRsMozos!mozo))
            .Tag = Trim(oRsMozos!Codigo)
        End With

        '    Set ItemM = Me.lvMozo.ListItems.Add(, , Trim(oRsMozos!mozo))
        '    ItemM.Tag = Trim(oRsMozos!Codigo)
        oRsMozos.MoveNext
    Loop

End Sub
Private Sub ConfigurarLV()
With Me.lvMozo
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "Mozo", 3700
    
End With
End Sub


Private Sub cmdAceptar_Click()
If Not Me.lvMozo.SelectedItem Is Nothing Then EnviaDatos

End Sub

Private Sub cmdAnterior_Click()
Desplazarse True
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdPosterior_Click()
Desplazarse False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
MsgBox KeyAscii
End Sub

Private Sub Form_Load()
InhabilitarCerrar Me
ConfigurarLV
CargarMozos
End Sub

Private Sub lvMozo_DblClick()
EnviaDatos
End Sub
