VERSION 5.00
Begin VB.Form Frmcalculo 
   Caption         =   "Determinar ..."
   ClientHeight    =   1416
   ClientLeft      =   6252
   ClientTop       =   3096
   ClientWidth     =   2784
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1416
   ScaleWidth      =   2784
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Tag             =   "9999"
      Top             =   0
      Width           =   2535
      Begin VB.TextBox tbolsas 
         Height          =   285
         Left            =   720
         TabIndex        =   1
         Tag             =   "9999"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox tconos 
         Height          =   285
         Left            =   720
         TabIndex        =   0
         Tag             =   "9999"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox tneto 
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Tag             =   "9999"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox tpc 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Tag             =   "9999"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox tpb 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Tag             =   "9999"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label c1 
         Caption         =   "Bolsas:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Tag             =   "9999"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label c1 
         Caption         =   "Conos:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Tag             =   "9999"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label c1 
         Caption         =   "Neto:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Tag             =   "9999"
         Top             =   960
         Width           =   615
      End
      Begin VB.Label c1 
         Caption         =   "x"
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   8
         Tag             =   "9999"
         Top             =   240
         Width           =   255
      End
      Begin VB.Label c1 
         Caption         =   "x"
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   7
         Tag             =   "9999"
         Top             =   600
         Width           =   255
      End
      Begin VB.Label cbruto 
         Alignment       =   1  'Right Justify
         Caption         =   "."
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Tag             =   "9999"
         Top             =   960
         Width           =   975
      End
   End
End
Attribute VB_Name = "Frmcalculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
tpc.Text = Format(GEN!GEN_TASA_LEG_VENC, "0.00")
tpb.Text = Format(GEN!GEN_TASA_LEG_ADE, "0.00")
End Sub

Private Sub tbolsas_GotFocus()
Azul tbolsas, tbolsas
End Sub

Private Sub tconos_GotFocus()
Azul tconos, tconos
End Sub

Private Sub tneto_Change()
LOC_BRUTO = CAL_BRUTO()
End Sub

Private Sub tneto_GotFocus()
Azul tneto, tneto
End Sub

Private Sub tneto_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
Unload Frmcalculo
End Sub

Private Sub tpb_Change()
LOC_BRUTO = CAL_BRUTO()
End Sub

Private Sub tpb_LostFocus()
If Val(tpb.Text) <> Val(GEN!GEN_TASA_LEG_ADE) Then
    GEN.Edit
    GEN!GEN_TASA_LEG_ADE = Val(tpb.Text)
    GEN.Update
    GEN.Requery
End If


End Sub

Private Sub tpc_Change()
LOC_BRUTO = CAL_BRUTO()
End Sub

Private Sub tpc_LostFocus()
If Val(tpc.Text) <> Val(GEN!GEN_TASA_LEG_VENC) Then
    GEN.Edit
    GEN!GEN_TASA_LEG_VENC = Val(tpc.Text)
    GEN.Update
    GEN.Requery
End If
End Sub

Private Sub tbolsas_Change()
LOC_BRUTO = CAL_BRUTO()
End Sub

Private Sub tbolsas_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 Unload Frmcalculo
 Exit Sub
End If

If KeyAscii <> 13 Then Exit Sub
LOC_BRUTO = CAL_BRUTO()
Azul tneto, tneto

End Sub

Private Sub tconos_Change()
LOC_BRUTO = CAL_BRUTO()
End Sub

Private Sub tconos_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 Unload Frmcalculo
 Exit Sub
End If

If KeyAscii <> 13 Then Exit Sub
Azul tbolsas, tbolsas
End Sub

Public Function CAL_BRUTO() As Currency
' CAMTEX
CAL_BRUTO = 0
Dim VAL_CONO As Currency
Dim VAL_BOLSA As Currency
VAL_CONO = Val(tconos.Text) * Val(tpc.Text)
VAL_BOLSA = Val(tbolsas.Text) * Val(tpb.Text)
CAL_BRUTO = Val(tneto.Text) - VAL_CONO - VAL_BOLSA
cbruto.Caption = Format(CAL_BRUTO, "0.00")

End Function

