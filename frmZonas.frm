VERSION 5.00
Begin VB.Form frmZonas 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zonas del Establecimiento"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13530
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmZonas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   13530
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   2760
      Picture         =   "frmZonas.frx":08CB
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdZonSig 
      Enabled         =   0   'False
      Height          =   495
      Left            =   1440
      Picture         =   "frmZonas.frx":0FB5
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdZonAnt 
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      Picture         =   "frmZonas.frx":16A0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label lblnZona 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Image imgZonas 
      BorderStyle     =   1  'Fixed Single
      Height          =   2160
      Index           =   0
      Left            =   50
      Picture         =   "frmZonas.frx":1D8B
      Stretch         =   -1  'True
      Top             =   50
      Visible         =   0   'False
      Width           =   3360
   End
End
Attribute VB_Name = "frmZonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vPagActZon As Integer
Private vPagTotZon As Integer
Private vIniLeft, vIniTop As Integer

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdZonAnt_Click()
Dim ini, fin, f As Integer
If vPagActZon = 2 Then
    ini = 1
    fin = 12
ElseIf vPagActZon = 1 Then
    Exit Sub
Else
    FF = vPagActZon - 1
    ini = (12 * FF) - 11
    fin = 12 * FF
End If

For f = ini To fin
    Me.imgZonas(f).Visible = True
    Me.lblnZona(f).Visible = True
Next
If vPagActZon > 1 Then
vPagActZon = vPagActZon - 1
    If vPagActZon = 1 Then: Me.cmdZonAnt.Enabled = False
    
    Me.cmdZonSig.Enabled = True
End If
End Sub

Private Sub cmdZonSig_Click()

    Dim ini, fin, f As Integer

    If vPagActZon = 1 Then
        ini = 1
        fin = 12
    ElseIf vPagActZon = vPagTotZon Then

        Exit Sub

    Else
        ini = (12 * vPagActZon) - 11
        fin = 12 * vPagActZon
    End If

    For f = ini To fin
        Me.imgZonas(f).Visible = False
        Me.lblnZona(f).Visible = False
    Next

    If vPagActZon < vPagTotZon Then
        vPagActZon = vPagActZon + 1

        If vPagActZon = vPagTotZon Then: Me.cmdZonSig.Enabled = False
    
        Me.cmdZonAnt.Enabled = True
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me

End Sub

Private Sub Form_Load()
    vIniLeft = 50
    vIniTop = 50
   CentrarFormulario MDIForm1, Me
    'Me.cmdZonAnt.Left = vIniLeft

    Dim VZONA    As Integer

    Dim oRsZonas As ADODB.Recordset

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpListarZonas"

    Set oRsZonas = oCmdEjec.Execute(, LK_CODCIA)

    Dim valor As Double

    Dim f     As Integer

    f = 1
    VZONA = oRsZonas.RecordCount
    vPagTotZon = oRsZonas.RecordCount / 12
    valor = VZONA / 12

    pos = InStr(Trim(str(valor)), ".")

    If pos <> 0 Then
        pos2 = Right(Trim(str(valor)), Len(Trim(str(valor))) - pos)
        ent = Left(Trim(str(valor)), pos - 1)

        If ent = "" Then: ent = 0

    End If

    'vPagActZon = vZona
    If VZONA > 0 Then vPagActZon = 1
    If VZONA > 12 Then Me.cmdZonSig.Enabled = True
    If pos2 > 0 Then: vPagTotZon = ent + 1

    Dim b As Integer
Dim xFILA As Integer
xFILA = 1
    For b = 1 To VZONA
        Load Me.imgZonas(b)
        Load Me.lblnZona(b)

        If f <= 4 Then '1° fila
            If f <> 1 Then
                vIniLeft = vIniLeft + Me.imgZonas(f - 1).Width
            End If

        ElseIf f <= 8 Then '2° fila
xFILA = 2
            If f = 5 Then
                vIniLeft = 50
                vIniTop = vIniTop + Me.imgZonas(f - 1).Height + Me.lblnZona(f - 1).Height
            Else
                vIniLeft = vIniLeft + Me.imgZonas(f - 1).Width
            End If

        ElseIf f <= 12 Then
xFILA = 3
            If f = 9 Then
                vIniLeft = 50
                vIniTop = vIniTop + Me.imgZonas(f - 1).Height + Me.lblnZona(f - 1).Height
            Else
                vIniLeft = vIniLeft + Me.imgZonas(f - 1).Width
            End If
        End If

        '        If f > 12 Then
        '            vIniLeft = 50
        '            f = 1
        '        ElseIf b > 1 Then
        '            vIniLeft = vIniLeft + Me.imgZonas(b - 1).Width
        '        End If
        '
        Me.imgZonas(b).Visible = True
        Me.imgZonas(b).Tag = oRsZonas!Codigo
        Me.imgZonas(b).Move vIniLeft, vIniTop
        '
        Me.lblnZona(b).Visible = True
        Me.lblnZona(b).Caption = Trim(oRsZonas!denomina)
        
        Me.lblnZona(b).Move vIniLeft, vIniTop + Me.imgZonas(b).Height  ' Me.imgZonas(b).Left, Me.imgZonas(b).Height
       
        If f = 12 Then
            f = 1
            vIniTop = 50
            vIniLeft = 50
        Else
            f = f + 1
        End If
        
        oRsZonas.MoveNext
    Next

    vIniLeft = 50
    vIniTop = 50
    
    '    Me.cmdZonAnt.Move vIniLeft, 2800 + 500
    '    Me.cmdZonSig.Move Me.cmdZonAnt.Width + vIniLeft, 2800 + 500
    '    Me.cmdSalir.Move Me.cmdZonAnt.Width + Me.cmdZonSig.Width + vIniLeft, 2800 + 500
    
    If xFILA = 1 Then
        Me.Height = Me.imgZonas(1).Height + Me.lblnZona(1).Height + 1000 + Me.cmdSalir.Height
        Me.cmdSalir.Top = Me.imgZonas(1).Height + Me.lblnZona(1).Height + Me.cmdSalir.Height
        Me.cmdZonAnt.Top = Me.imgZonas(1).Height + Me.lblnZona(1).Height + Me.cmdSalir.Height
        Me.cmdZonSig.Top = Me.imgZonas(1).Height + Me.lblnZona(1).Height + Me.cmdSalir.Height
        ElseIf xFILA = 2 Then
        Me.Height = (Me.imgZonas(1).Height * 2) + (Me.lblnZona(1).Height * 2) + 1000 + (Me.cmdSalir.Height * 2)
         Me.cmdSalir.Top = (Me.imgZonas(1).Height * 2) + (Me.lblnZona(1).Height * 2) + (Me.cmdSalir.Height * 2)
        Me.cmdZonAnt.Top = (Me.imgZonas(1).Height * 2) + (Me.lblnZona(1).Height * 2) + (Me.cmdSalir.Height * 2)
        Me.cmdZonSig.Top = (Me.imgZonas(1).Height * 2) + (Me.lblnZona(1).Height * 2) + (Me.cmdSalir.Height * 2)
    End If
End Sub

Private Sub imgZonas_Click(Index As Integer)
'Dim oFrmUbi As New frmPrincipal
frmDisMesas.VZONA = Me.imgZonas(Index).Tag
frmDisMesas.Caption = "Distribución de Mesas : " & lblnZona(Index).Caption
frmDisMesas.Show
Unload Me
End Sub


