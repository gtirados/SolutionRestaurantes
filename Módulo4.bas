Attribute VB_Name = "Módulo4"
Public VReporte As New CRAXDRT.Report
Public Enum Valores
    InicializarFormulario
    Nuevo
    grabar
    cancelar
    Editar
    buscar
    AntesDeActualizar
    Eliminar           'LINEA NUEVA
End Enum

Public Declare Function DrawMenuBar Lib "User32" _
      (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "User32" _
      (ByVal hMenu As Long) As Long
Public Declare Function GetSystemMenu Lib "User32" _
        (ByVal hwnd As Long, _
        ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "User32" _
        (ByVal hMenu As Long, _
        ByVal nPosition As Long, _
        ByVal wFlags As Long) As Long
        
Public oCmdEjec As New ADODB.Command

Public Sub LimpiaParametros(oCmd As ADODB.Command)
    oCmd.ActiveConnection = Pub_ConnAdo
    oCmd.CommandType = adCmdStoredProc
    Pub_ConnAdo.CursorLocation = adUseClient

    For i = oCmd.Parameters.count - 1 To 0 Step -1
        oCmd.Parameters.Delete i
    Next

End Sub

Public Sub InhabilitarCerrar(ofrm As Form)
Dim hMenu As Long
Dim menuItemCount As Long
'Obtenemos un handle al menú de sistema del formulario
hMenu = GetSystemMenu(ofrm.hwnd, 0)
If hMenu Then
    'Obtenemos el número de elementos del menú
    menuItemCount = GetMenuItemCount(hMenu)
    'Eliminamos el elemento Cerrar, que es el último
    'Los elemento empiezan a numerarse en cero por lo que el
    'último es menuItemCount - 1
     Call RemoveMenu(hMenu, menuItemCount - 1, _
                      MF_REMOVE Or MF_BYPOSITION)
    'Eliminamos la barra de separación que hay justo antes de la opción Cerrar
    Call RemoveMenu(hMenu, menuItemCount - 2, _
                      MF_REMOVE Or MF_BYPOSITION)
    'Forzamos el redibujado del menú. Esto refresca la barra de título
    'y deja la X deshabilitada
    Call DrawMenuBar(ofrm.hwnd)
End If
End Sub


Public Sub MostrarErrores(xError As ErrObject)
MsgBox "Descripcion del Error: " & xError.Description & vbCrLf & _
"Origen del Error: " & xError.Source & vbCrLf & "Número de Error: " & xError.Number, vbCritical, NombreProyecto
End Sub


Public Sub LimpiarControles(Frm As Form)
   Dim i
   For i = 0 To Frm.Controls.count - 1
      If TypeOf Frm.Controls(i) Is TextBox Then
         Frm.Controls(i).Text = ""
      ElseIf TypeOf Frm.Controls(i) Is label And Frm.Controls(i).Tag = "X" Then
          Frm.Controls(i).Caption = ""
      ElseIf TypeOf Frm.Controls(i) Is ComboBox Then
Frm.Controls(i).ListIndex = -1
      End If
   Next i
End Sub

Public Sub ActivarControles(Frm As Form)
Dim j
For j = 0 To Frm.Controls.count - 1
    If TypeOf Frm.Controls(j) Is TextBox Then
        Frm.Controls(j).Enabled = True
    End If
    If TypeOf Frm.Controls(j) Is DataCombo Then
        Frm.Controls(j).Enabled = True
    End If
    If TypeOf Frm.Controls(j) Is ComboBox Then
        Frm.Controls(j).Enabled = True
    End If
     If TypeOf Frm.Controls(j) Is DTPicker Then
        Frm.Controls(j).Enabled = True
    End If
Next
End Sub

Public Function Mayusculas(Caracter As Integer) As Integer
    'Para escribir en Mayusculas
    
    Mayusculas = Asc(UCase(Chr(Caracter)))
End Function

Public Sub DesactivarControles(Frm As Form)
Dim j
For j = 0 To Frm.Controls.count - 1
    If TypeOf Frm.Controls(j) Is TextBox And Frm.Controls(j).Tag = "X" Then
        Frm.Controls(j).Enabled = False
    End If
    If TypeOf Frm.Controls(j) Is DataCombo Then
        Frm.Controls(j).Enabled = False
    End If
    If TypeOf Frm.Controls(j) Is ComboBox Then
        Frm.Controls(j).Enabled = False
    End If
    If TypeOf Frm.Controls(j) Is DTPicker Then
        Frm.Controls(j).Enabled = False
    End If
Next
End Sub

