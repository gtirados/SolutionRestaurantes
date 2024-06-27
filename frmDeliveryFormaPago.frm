VERSION 5.00
Begin VB.Form frmDeliveryFormaPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forma de Pago"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
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
   ScaleHeight     =   2070
   ScaleWidth      =   5490
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   720
      Left            =   3840
      Picture         =   "frmDeliveryFormaPago.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   720
      Left            =   3840
      Picture         =   "frmDeliveryFormaPago.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtAbono 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblVuelto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SU VUELTO:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   4
      Top             =   1627
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ABONA CON:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   435
      TabIndex        =   2
      Top             =   1020
      Width           =   1260
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   945
      TabIndex        =   0
      Top             =   307
      Width           =   750
   End
End
Attribute VB_Name = "frmDeliveryFormaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gTOTAL As Double
Private vPUNTO As Boolean

Private Sub cmdAceptar_Click()

    If Not IsNumeric(Me.txtAbono.Text) Then
        MsgBox "El Abotno no es correcto.", vbCritical, Pub_Titulo

        Exit Sub

    End If
    
    If val(Me.lblVuelto.Caption) < 0 Then
        MsgBox "El Abono es incorrecto.", vbCritical, Pub_Titulo
        Me.txtAbono.SetFocus
        Me.txtAbono.SelStart = 0
        Me.txtAbono.SelLength = Len(Me.txtAbono.Text)
        Exit Sub
    End If
    
    If val(Me.txtAbono.Text) <= 0 Then
        MsgBox "El Abono es incorrecto.", vbCritical, Pub_Titulo
        Me.txtAbono.SetFocus
        Me.txtAbono.SelStart = 0
        Me.txtAbono.SelLength = Len(Me.txtAbono.Text)
        Exit Sub
    End If
    
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_DELIVERY_ASIGNAR_FORMAPAGO"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim xEXITO As String

    xEXITO = ""
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adVarChar, adParamInput, 3, frmDeliveryApp.lblserie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adBigInt, adParamInput, , frmDeliveryApp.lblnumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PAGO", adDouble, adParamInput, , Me.txtAbono.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@EXITO", adVarChar, adParamOutput, 300, xEXITO)
    oCmdEjec.Execute
    
    xEXITO = oCmdEjec.Parameters("@EXITO").Value

    If Len(Trim(xEXITO)) = 0 Then
        MsgBox "Datos almacenados Correctamente.", vbInformation, Pub_Titulo
        Unload Me
    Else
        MsgBox xEXITO, vbCritical, Pub_Titulo
    End If

End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Me.lblTotal.Tag = gTOTAL
    Me.lblTotal.Caption = "S/. " + Format(gTOTAL, "#####.00")
   
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_DELIVERY_DATOS_CLIENTE"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim oRSp As ADODB.Recordset

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numser", adVarChar, adParamInput, 3, frmDeliveryApp.lblserie.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numfac", adBigInt, adParamInput, , frmDeliveryApp.lblnumero.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    Set oRSp = oCmdEjec.Execute

    If Not oRSp.EOF Then Me.txtAbono.Text = oRSp!pago
    
 'Me.txtAbono.Text = "0.00"
    Me.txtAbono.SelStart = 0
    Me.txtAbono.SelLength = Len(Me.txtAbono.Text)
    If val(Me.txtAbono.Text) = 0 Then Me.lblVuelto.Caption = "0.00"

End Sub

Private Sub txtAbono_Change()

    If InStr(Me.txtAbono.Text, ".") Then
        vPUNTO = True
    Else
        vPUNTO = False
    End If

    If IsNumeric(Me.txtAbono.Text) Then
        Me.lblVuelto.Caption = val(Me.txtAbono.Text) - val(Me.lblTotal.Tag)
    Else
        Me.lblVuelto.Caption = "0.00"
    End If

End Sub

Private Sub txtAbono_KeyPress(KeyAscii As Integer)

    If NumerosyPunto(KeyAscii) Then KeyAscii = 0
    If KeyAscii = 46 Then
        If vPUNTO Or Len(Trim(Me.txtAbono.Text)) = 0 Then
            KeyAscii = 0
        End If
    End If

    If KeyAscii = vbKeyReturn Then cmdAceptar_Click
End Sub
