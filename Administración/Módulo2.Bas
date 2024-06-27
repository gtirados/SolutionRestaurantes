Attribute VB_Name = "Módulo2"
Public NUMERO As Integer
Public CAMPOS1 As Integer
Public exito As Boolean

Public Sub Azul(vart As Variant, varc As TextBox)
If varc.Enabled = True And varc.Visible = True Then
   varc.SetFocus
   
   varc.SelStart = 0
   varc.SelLength = Len(vart)
End If
End Sub
Public Sub Azul2(vart As Variant, varc As MaskEdBox)
If varc.Enabled = True And varc.Visible = True Then
   varc.SetFocus
   varc.SelStart = 0
   varc.SelLength = Len(vart)
End If
End Sub
Public Sub Azul3(vart As Variant, varc As RichTextBox)
If varc.Enabled And varc.Visible Then
   varc.SetFocus
   varc.SelStart = 0
   varc.SelLength = Len(vart)
End If
End Sub

Public Function ALINEA(VAR As String) As String
Dim TEMP As String * 15
Dim N1 As Integer
Dim N2 As Integer
N1 = InStr(1, VAR, " ") - 1
N2 = Len(VAR) - N1
VAR = String(N2, "    ") + Left(VAR, N1)
ALINEA = VAR
End Function

Public Function Nulo_Valor0(Optional valor) As Variant
If IsNull(valor) = True Or valor = "" Then
   Nulo_Valor0 = 0
Else
   Nulo_Valor0 = valor
End If

End Function
Public Function Nulo_Valors(Optional valor) As Variant
If IsNull(valor) = True Then
   Nulo_Valors = ""
Else
   Nulo_Valors = valor
End If

End Function

Public Function NULO_DATE(Optional FEC) As Date
If IsDate(FEC) Then
   NULO_DATE = CDate(FEC)
End If

End Function
Public Function redondea(valor As Variant) As Variant
redondea = Format(valor, "########0.00")
End Function
Public Sub SOLO_DECIMAL(wsTexto As TextBox, Optional wsKeyAscii)
Dim car
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If wsKeyAscii = 45 Then
      If wsTexto.text <> "" Then
         Beep
         wsKeyAscii = 0
         Exit Sub
      End If
    End If
    If wsKeyAscii = 46 Then
      If InStr(1, wsTexto.text, ".") <> 0 Then
        Beep
        wsKeyAscii = 0
        Exit Sub
      End If
    End If
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And car <> "." And car <> "-" Then
          wsKeyAscii = 0
          Beep
          Exit Sub
        End If
    End If
End Sub
Public Sub SOLO_ENTERO(Optional tecla)

Dim car As String, Longt As Integer
car = Chr$(tecla)
car = UCase$(Chr$(tecla))
tecla = Asc(car)
If car < "0" Or car > "9" Then
    If tecla <> 8 And tecla <> 13 Then
        tecla = 0
        Beep
    End If
End If
End Sub

Public Sub BORRA_FIELDS(NUM As Integer, ARRAYY As Variant)
Dim X As Integer
For X = 1 To NUM
 ARRAYY(X).text = ""
Next

End Sub

