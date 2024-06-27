VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmBusProd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Busqueda"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8865
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
   ScaleHeight     =   3585
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvData 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   5106
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtBuscar 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   7095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DESCRIPCIÓN:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmBusProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vTipo As String
Private ors As ADODB.Recordset
Private pos As Integer
Private PresEnter As Boolean

Private Sub ConfiguraLV()
With Me.lvData
    .Gridlines = True
    .LabelEdit = lvwManual
    .FullRowSelect = True
    .View = lvwReport
    .ColumnHeaders.Add , , "Código"
    .ColumnHeaders.Add , , "Descripción", 5000
    .ColumnHeaders.Add , , "Unidad", 1000
    .HideSelection = False
End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
ConfiguraLV
PresEnter = False
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SpListarProdxTipo"
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Tipo", adChar, adParamInput, 1, vTipo)
Set ors = oCmdEjec.Execute

End Sub

Private Sub txtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
If Me.lvData.ListItems.count > 0 Then
    If KeyCode = vbKeyDown Then
        If pos < Me.lvData.ListItems.count Then
        pos = pos + 1
        Me.lvData.ListItems(pos).Selected = True
        End If
    ElseIf KeyCode = vbKeyUp Then
        If pos <> 1 Then
        pos = pos - 1
        Me.lvData.ListItems(pos).Selected = True
        End If
    End If
End If
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
 Dim dd As Fields
            Dim itemL As Object
    If PresEnter Then
        If Not Me.lvData.SelectedItem Is Nothing Then
            If Me.txtBuscar.Tag <> Me.txtBuscar.Text Then
                PresEnter = True
                If Len(Trim(Me.txtBuscar.Text)) <> 0 Then
                    ors.Filter = "Descripcion like '%" & Me.txtBuscar.Text & "%'"
                    Me.txtBuscar.Tag = Me.txtBuscar.Text
                    Me.lvData.ListItems.Clear
                    
                    pos = 1
                    Do While Not ors.EOF
                        Set itemL = Me.lvData.ListItems.Add(, , ors!Codigo)
                        itemL.SubItems(1) = Trim(ors!DESCRIPCION)
                        itemL.SubItems(2) = Trim(ors!UNIDAD)
                        ors.MoveNext
                    Loop
                End If
            Else
                Dim ofrmAsigCant As New frmAsigCantidad
                ofrmAsigCant.xTipo = vTipo
                If frmListaProd.cboTipoProd.ListIndex = 1 Then 'materia prima
                
                ofrmAsigCant.Caption = "Asignar Precio de Mercado"
                ofrmAsigCant.Label1.Caption = "Precio"
                End If
                ofrmAsigCant.Show vbModal
            End If
        End If
    Else
        
        If Len(Trim(Me.txtBuscar.Text)) <> 0 Then
            ors.Filter = "Descripcion like '%" & Me.txtBuscar.Text & "%'"
            Me.txtBuscar.Tag = Me.txtBuscar.Text
            Me.lvData.ListItems.Clear
            'Dim dd As Fields
            'Dim itemL As Object
            
            If Not ors.EOF Then
                pos = 1
                Do While Not ors.EOF
                    Set itemL = Me.lvData.ListItems.Add(, , ors!Codigo)
                    itemL.SubItems(1) = Trim(ors!DESCRIPCION)
                    itemL.SubItems(2) = Trim(ors!UNIDAD)
                    ors.MoveNext
                Loop
                PresEnter = True
            End If
            
        End If
    End If
    
    
End If
End Sub
