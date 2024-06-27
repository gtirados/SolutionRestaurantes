VERSION 5.00
Begin VB.Form frmComandaMotivoElimina 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Motivos de Anulación"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12255
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   12255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   840
      Left            =   8760
      TabIndex        =   5
      Top             =   6000
      Width           =   1695
   End
   Begin VB.TextBox txtOtros 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   4
      Top             =   6360
      Width           =   8535
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   840
      Left            =   10500
      TabIndex        =   3
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdMotivo 
      Height          =   1440
      Index           =   0
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdNext 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   9780
      Picture         =   "frmComandaMotivoElimina.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton cmdPrev 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   120
      Picture         =   "frmComandaMotivoElimina.frx":4F0A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Otros:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   6120
      Width           =   540
   End
End
Attribute VB_Name = "frmComandaMotivoElimina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vPAGINATOTAL, vPAGINACTUAL As Integer
Public gAcepta As Boolean
Public gIDmotivo As Integer
Public gMOTIVO As String

Private Sub CargaMotivos()

    Dim orsM As ADODB.Recordset

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_MOTIVOS_LISTADO"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Execute
    Set orsM = oCmdEjec.Execute

    Me.cmdPrev.Enabled = False

    If Me.cmdMotivo.count <> 1 Then

        For i = 1 To Me.cmdMotivo.count - 1
            Unload Me.cmdMotivo(i)
        Next

    End If
    
    Dim cont    As Integer

    Dim posleft As Integer, postop As Integer

    Dim fila    As Integer

    fila = 1
    posleft = 120
    postop = 120
    
    For cont = 1 To orsM.RecordCount
        Load Me.cmdMotivo(cont)
  
        If fila <= 4 Then '1° fila
            If fila = 1 Then
                posleft = posleft + Me.cmdPrev.Width
            Else
                posleft = posleft + Me.cmdMotivo(fila - 1).Width
            End If

            ' MsgBox "james"
        ElseIf fila <= 9 Then '2° fila
          
            If fila = 5 Then
                posleft = 120
                postop = postop + Me.cmdMotivo(fila - 1).Height
            Else
                posleft = posleft + Me.cmdMotivo(fila - 1).Width
            End If

        ElseIf fila <= 14 Then

            If fila = 10 Then
                posleft = 120
                postop = postop + Me.cmdMotivo(fila - 1).Height
            Else
                posleft = posleft + Me.cmdMotivo(fila - 1).Width
            End If

        ElseIf fila <= 18 Then

            If fila = 15 Then
                posleft = 120
                postop = postop + Me.cmdMotivo(fila - 1).Height
            Else
                posleft = posleft + Me.cmdMotivo(fila - 1).Width
            End If
        End If

        Me.cmdMotivo(cont).Left = posleft
        cmdMotivo(cont).Top = postop
        
        cmdMotivo(cont).Visible = True
        cmdMotivo(cont).Tag = orsM!IDE
        
        cmdMotivo(cont).Caption = orsM!MOT

        If fila = 18 Then
            fila = 1
            postop = 120
            posleft = 120
        Else
            fila = fila + 1
        End If

        orsM.MoveNext
    Next
    
    Dim valor As Double

    valor = orsM.RecordCount / 13
    pos = InStr(Trim(str(valor)), ".")
    
    If pos <> 0 Then
        If pos = 1 Then
            ent = Left(CStr(valor), pos)
        Else
            ent = Left(CStr(valor), pos - 1)
        End If

    Else
        ent = Int(valor)
    End If
    
    If pos <> 0 Then
        pos2 = Right(Trim(str(valor)), Len(Trim(str(valor))) - pos)
    Else
        pos2 = 0
    End If

    If pos2 > 0 Then
        ent = ent + 1
    End If

    vPAGINATOTAL = ent

    If valor <> 0 Then vPAGINACTUAL = 1

    If orsM.RecordCount > 13 Then: Me.cmdNext.Enabled = True

End Sub

Private Sub cmdAceptar_Click()

    If Len(Trim(Me.txtOtros.Text)) = 0 Then
        MsgBox "Debe ingresar el motivo de Eliminación.", vbInformation, Pub_Titulo
        Me.txtOtros.SetFocus
    Else

        If MsgBox("¿Desea continuar con la operación.?", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub
        gAcepta = True
        gIDmotivo = -1
        gMOTIVO = Me.txtOtros.Text
        Unload Me
    End If

End Sub

Private Sub cmdCancelar_Click()
gAcepta = False
Unload Me
End Sub

Private Sub cmdMotivo_Click(Index As Integer)

    If MsgBox("¿Desea continuar con la operación.?", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub
    gAcepta = True
    gIDmotivo = Index
    Unload Me
End Sub

Private Sub cmdNext_Click()

    Dim ini, fin, f As Integer

    If vPagActual = 1 Then
        ini = 1
        fin = 18
    ElseIf vPAGINACTUAL = vPAGINATOTAL Then

        Exit Sub

    Else
        ini = (18 * vPAGINACTUAL) - 17
        fin = 18 * vPAGINACTUAL
    End If

    For f = ini To fin
        Me.cmdMotivo(f).Visible = False
    Next

    If vPAGINACTUAL < vPAGINATOTAL Then
        vPAGINACTUAL = vPAGINACTUAL + 1

        If vPAGINACTUAL = vPAGINATOTAL Then: Me.cmdNext.Enabled = False
    
        Me.cmdPrev.Enabled = True
    End If

End Sub

Private Sub cmdPrev_Click()
CargaMotivos
   Dim ini, fin, f As Integer

    If vPAGINACTUAL = 2 Then
        ini = 1
        fin = 18
    ElseIf vPAGINACTUAL = 1 Then

        Exit Sub

    Else
        FF = vPAGINACTUAL - 1
        ini = (18 * FF) - 17
        fin = 18 * FF
    End If

    For f = ini To fin
        Me.cmdMotivo(f).Visible = True
    Next

    If vPAGINACTUAL > 1 Then
        vPAGINACTUAL = vPAGINACTUAL - 1

        If vPAGINACTUAL = 1 Then: Me.cmdPrev.Enabled = False

        Me.cmdNext.Enabled = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
CargaMotivos
gAcepta = False
End Sub
