VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmResumen 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resumen"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7695
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
   ScaleHeight     =   5370
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvResumen 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9128
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmResumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ConfigurarLv()
With Me.lvResumen
    .ColumnHeaders.Add , , "Plato", 6000
    .ColumnHeaders.Add , , "Cant.", 1000, 1
    .FullRowSelect = True
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
ConfigurarLv
Dim fila As Integer
Dim Filab As Integer
Dim ve As Boolean
v = False
For fila = 1 To frmComanda.lvPlatos.ListItems.count
ve = False
    If Me.lvResumen.ListItems.count = 0 Then
        With Me.lvResumen.ListItems.Add(, , frmComanda.lvPlatos.ListItems(fila).Text)
        .SubItems(1) = frmComanda.lvPlatos.ListItems(fila).SubItems(3)
        .Tag = frmComanda.lvPlatos.ListItems(fila).Tag
        End With
    Else
        For Filab = 1 To Me.lvResumen.ListItems.count
            If CStr(Me.lvResumen.ListItems(Filab).Tag) = frmComanda.lvPlatos.ListItems(fila).Tag Then
                ve = True
                Exit For
            End If
        Next
        If ve Then
            Me.lvResumen.ListItems(Filab).SubItems(1) = FormatNumber(CDec(Me.lvResumen.ListItems(Filab).SubItems(1)) + CDec(frmComanda.lvPlatos.ListItems(fila).SubItems(3)), 2) 'linea nueva 04-08-2011
        Else
         With Me.lvResumen.ListItems.Add(, , frmComanda.lvPlatos.ListItems(fila).Text)
        .SubItems(1) = frmComanda.lvPlatos.ListItems(fila).SubItems(3)
        .Tag = frmComanda.lvPlatos.ListItems(fila).Tag
        End With
        End If
    End If
Next
End Sub
