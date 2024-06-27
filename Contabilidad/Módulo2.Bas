Attribute VB_Name = "Módulo2"
Public CAMPOS1 As Integer
Public exito As Boolean
Public Declare Function SendMessageLong Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Res As Long

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

Public Function ALINEA(var As String) As String
Dim TEMP As String * 15
Dim N1 As Integer
Dim N2 As Integer
    N1 = InStr(1, var, " ") - 1
    N2 = Len(var) - N1
    var = String(N2, "    ") + Left(var, N1)
    ALINEA = var
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
      If wsTexto.Text <> "" Then
         Beep
         wsKeyAscii = 0
         Exit Sub
      End If
    End If
    If wsKeyAscii = 46 Then
      If InStr(1, wsTexto.Text, ".") <> 0 Then
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

Public Sub LlenaCboTablas(ByVal cbo As ComboBox, ByVal iTipReg As Integer, sCodCia)
    SQ_OPER = 2
    PUB_TIPREG = iTipReg
    PUB_CODCIA = sCodCia
    LEER_TAB_LLAVE
    Do While Not tab_mayor.EOF
        cbo.AddItem tab_mayor("Tab_nomlargo") & Space(50) & tab_mayor("tab_numtab")
        tab_mayor.MoveNext
    Loop
End Sub
