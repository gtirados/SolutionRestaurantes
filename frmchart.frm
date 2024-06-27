VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmchart 
   BackColor       =   &H00F5F1EC&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12870
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdview 
      BackColor       =   &H00F5F1EC&
      Height          =   570
      Left            =   11910
      Picture         =   "frmchart.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4905
      Width           =   810
   End
   Begin MSChart20Lib.MSChart Grafico 
      Height          =   4275
      Left            =   285
      OleObjectBlob   =   "frmchart.frx":0DAA
      TabIndex        =   0
      Top             =   585
      Width           =   12435
   End
   Begin MSMask.MaskEdBox txtCampo1 
      Height          =   315
      Left            =   8310
      TabIndex        =   2
      Tag             =   "9999"
      Top             =   5070
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   14737632
      ForeColor       =   128
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtcampo2 
      Height          =   330
      Left            =   10515
      TabIndex        =   3
      Tag             =   "9999"
      Top             =   5055
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   14737632
      ForeColor       =   128
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label lblUnidad 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009C3000&
      Height          =   195
      Left            =   345
      TabIndex        =   7
      Top             =   5010
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Fin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009C3000&
      Height          =   195
      Index           =   2
      Left            =   9615
      TabIndex        =   6
      Top             =   5130
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Ini."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009C3000&
      Height          =   195
      Index           =   1
      Left            =   7425
      TabIndex        =   5
      Top             =   5100
      Width           =   810
   End
   Begin VB.Label lblDescripcion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   285
      TabIndex        =   1
      Top             =   180
      Width           =   12405
   End
End
Attribute VB_Name = "frmchart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RSOrigen As New ADODB.Recordset

Public Sub ArtiFacart()
Dim Matrix() As String
Dim iRows As Integer
Dim iRow As Integer
Dim iRecorrido  As Integer
Dim NroMeses As Integer
Dim iDecMes As Integer
    
    NroMeses = DateDiff("m", txtCampo1.Text, txtcampo2.Text)
    ReDim Matrix(1 To NroMeses, 1 To 3)
    For iRecorrido = NroMeses To 1 Step -1
        Matrix(iRecorrido, 1) = Left(Format(DateAdd("m", iDecMes, txtcampo2.Text), "MMMM"), 3) & "-" & Format(DateAdd("m", iDecMes, txtcampo2.Text), "YYYY")
        iDecMes = iDecMes - 1
    Next iRecorrido
    SQ_OPER = 1
    PUB_KEY = PUB_CODART
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If Not art_LLAVE.EOF Then
        lblDescripcion.Caption = art_LLAVE("art_nombre")
    End If
    PUB_SECUEN = 0
    LEER_PRE_LLAVE
    If Not pre_llave.EOF Then
        lblUnidad.Caption = "Unidad = " & pre_llave("pre_unidad")
    End If
    
    archi = "SELECT MONTH(FAR_FECHA) AS MES, SUM(FAR_CANTIDAD) AS Cantidad, Year(FAR_FECHA) AS Anio , FAR_TIPMOV " ', FAR_CODART, FAR_CODCIA
    archi = archi & "From FACART WHERE (FAR_TIPMOV = 10 or FAR_TIPMOV = 20) AND FAR_ESTADO <> 'E' AND FAR_CODART = " & PUB_CODART
    archi = archi & " AND FAR_FECHA_COMPRA >= '" & txtCampo1.Text & "' AND FAR_FECHA_COMPRA <= '" & txtcampo2.Text & "' "
    archi = archi & " GROUP BY MONTH(FAR_FECHA), FAR_TIPMOV, FAR_CODART, Year(FAR_FECHA), FAR_CODCIA "
            
    Set RSOrigen = OpenSQLForwardOnly(archi)
    If Not RSOrigen.EOF Then
        iRows = RSOrigen.RecordCount
        
        Do While Not RSOrigen.EOF
            iRow = iRow + 1
            For iRecorrido = 1 To NroMeses
                If Matrix(iRecorrido, 1) = Left(Format("01/" & RSOrigen("Mes") & "/2000", "MMMM"), 3) & "-" & RSOrigen("Anio") Then
                    If RSOrigen("Far_tipmov") = 10 Then
                        Matrix(iRecorrido, 3) = RSOrigen("Cantidad")
                    ElseIf RSOrigen("Far_tipmov") = 20 Then
                        Matrix(iRecorrido, 2) = RSOrigen("Cantidad")
                    End If
                End If
            Next
            RSOrigen.MoveNext
        Loop
        'Set Grafico.DataSource = RSOrigen
        Grafico.ChartData = Matrix
    Else
        MsgBox "No existen Movimientos", vbInformation, Pub_Titulo
        Unload Me
        GoTo SALIR
    End If
    
    RSOrigen.Close
    Set RSOrigen = Nothing
    Grafico.Column = 1
    Grafico.ColumnLabel = "Compras"
    Grafico.Column = 2
    Grafico.ColumnLabel = "Ventas"
    frmchart.Show 1
SALIR:
End Sub

Private Sub cmdview_Click()
    ArtiFacart
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    txtcampo2.Text = LK_FECHA_DIA
    txtCampo1.Text = DateAdd("m", -3, LK_FECHA_DIA)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RSOrigen = Nothing
    Set frmchart = Nothing
End Sub

Private Sub Grafico_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub
