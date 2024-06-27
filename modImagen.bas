Attribute VB_Name = "modImagen"
Public Function Imagen_Array(Path_Imagen As String) As ADODB.Stream
  
Dim rs As Recordset
Dim Stream As ADODB.Stream
  

On Error GoTo Error_Sub
  
    
    'Nuevo objeto ADODB Stream
    Set Stream = New ADODB.Stream
      
    ' dato de tipo binario
    Stream.Type = adTypeBinary
      
    Stream.Open
        ' verifica que la ruta del gráfico no sea una cadena vacía
        If Len(Path_Imagen) <> 0 Then
              
            ' lee la imagen desde el path
            Stream.LoadFromFile Path_Imagen
         
            
             'Stream.Read
           
        End If
    ' cierra el recordset y elimina la referencia
  
    
    ' Retorno
    Set Imagen_Array = Stream
'      If Stream.State = adStateOpen Then
'        Stream.Close
'    End If
'    If Not Stream Is Nothing Then
'        Set Stream = Nothing
'    End If
Exit Function
  
Error_Sub:
  If Err.Number <> 0 Then
    MsgBox CStr(Err) & "  " & Error, vbExclamation
  End If
    
End Function

