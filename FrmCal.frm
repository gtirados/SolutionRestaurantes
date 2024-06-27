VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCal 
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   2910
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   2910
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox textovar 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid gr 
      Height          =   2415
      Left            =   975
      TabIndex        =   0
      Top             =   1110
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   4260
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   9128212
      ScrollBars      =   2
      Appearance      =   0
   End
   Begin VB.TextBox t_vuelto 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox t_total 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   960
      X2              =   2880
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   2880
      Y1              =   3900
      Y2              =   3900
   End
   Begin VB.Label importer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label lop 
      AutoSize        =   -1  'True
      Caption         =   "Vuelto :"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   4080
      Width           =   540
   End
   Begin VB.Label lop 
      AutoSize        =   -1  'True
      Caption         =   "Total: "
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   8
      Top             =   3600
      Width           =   450
   End
   Begin VB.Label lop 
      AutoSize        =   -1  'True
      Caption         =   "Recibido :"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label lop 
      AutoSize        =   -1  'True
      Caption         =   "Total Venta :"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pulsar F1 para Retornar..."
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label lbldocu 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   2880
   End
End
Attribute VB_Name = "FrmCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
Unload FrmCal
End If
End Sub

Private Sub Form_Load()
  FrmCal.Left = (Screen.Width - FrmCal.Width) / 2
  FrmCal.Top = (Screen.Height - FrmCal.Height) / 2
  gr.Cols = 2
  gr.Rows = 2
  gr.ColWidth(0) = 300
  gr.ColWidth(1) = 1000
  gr.TextMatrix(0, 0) = "+/-"
  gr.TextMatrix(0, 1) = "Importe"
  
  
  
  

End Sub

Private Sub gr_Click()
TEXTOVAR.Left = gr.Left + gr.CellLeft
TEXTOVAR.Width = gr.CellWidth
TEXTOVAR.Height = gr.CellHeight
TEXTOVAR.Top = gr.Top + gr.CellTop
TEXTOVAR.Tag = gr.TextMatrix(gr.Row, gr.COL)
TEXTOVAR.Text = gr.TextMatrix(gr.Row, gr.COL)
End Sub

Private Sub gr_EnterCell()
TEXTOVAR.Left = gr.Left + gr.CellLeft
TEXTOVAR.Width = gr.CellWidth
TEXTOVAR.Height = gr.CellHeight
TEXTOVAR.Top = gr.Top + gr.CellTop
TEXTOVAR.Tag = gr.TextMatrix(gr.Row, gr.COL)
TEXTOVAR.Text = gr.TextMatrix(gr.Row, gr.COL)

End Sub

Private Sub gr_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Exit Sub
If gr.COL = 0 Then
  'textovar.MaxLength = 1
  If KeyAscii = 45 Or KeyAscii = 43 Then
  Else
    TEXTOVAR.Visible = False
    KeyAscii = 0
    Exit Sub
  End If
Else
'textovar.MaxLength = 0
End If
TEXTOVAR.Visible = True
If KeyAscii <> 13 Then TEXTOVAR.Text = Chr(KeyAscii) 'gr.TextMatrix(gr.Row, gr.Col)
TEXTOVAR.SelStart = Len(TEXTOVAR.Text)
TEXTOVAR.SetFocus
If gr.COL = 0 Then
    If KeyAscii = 45 Or KeyAscii = 43 Then
     textovar_KeyPress 13
    End If
End If

End Sub

Private Sub Label3_Click()
End Sub

Private Sub textovar_Change()
gr.TextMatrix(gr.Row, gr.COL) = TEXTOVAR.Text
calcu
End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
gr.TextMatrix(gr.Row, gr.COL) = TEXTOVAR.Tag
TEXTOVAR.Text = TEXTOVAR.Tag
TEXTOVAR.Visible = False
gr.SetFocus
Exit Sub
End If

If KeyAscii = 13 Then
'gr.TextMatrix(gr.Row, gr.Col) = textovar.Text
 If gr.COL = 1 And (gr.Row = gr.Rows - 1) Then
   gr.Rows = gr.Rows + 1
   gr.Row = gr.Row + 1
   gr.COL = 0
   GoTo sa
 ElseIf gr.COL = 1 Then
   gr.Row = gr.Row + 1
   gr.COL = 0
   GoTo sa
 ElseIf gr.COL = 0 Then
   gr.COL = 1
   GoTo sa
 End If
sa:
 TEXTOVAR.Visible = False
 gr.SetFocus
End If

End Sub

Public Sub calcu()
Dim I As Integer
Dim wtotal As Currency
Dim wsi As Integer
wtotal = 0
wsi = 0
For I = 1 To gr.Rows - 1
If gr.TextMatrix(I, 0) = "-" Then
 wsi = -1
Else
 wsi = 1
End If
wtotal = wtotal + (wsi * Val(gr.TextMatrix(I, 1)))
Next I
If (Val(Format(t_total.Text, "0.00")) - wtotal) < 0 Then
 t_vuelto.ForeColor = QBColor(12)
Else
t_vuelto.ForeColor = QBColor(0)
End If
t_vuelto.Text = Format(Val(Format(t_total.Text, "0.00")) - wtotal, "#,##0.00")
importer.Caption = Format(wtotal, "#,##0.00")
End Sub
