Sub CopiarDatos2()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim rangoCopiar As Range
    Dim celdaDestino As Range
    
    ' Definir las hojas
    Set wsOrigen = ThisWorkbook.Sheets("MAR.NOV.")
    Set wsDestino = ThisWorkbook.Sheets("MAR.DIC.")
    
    ' Definir el rango a copiar
    Set rangoCopiar = wsOrigen.Range("H12:H63") ' Cambia este rango según tus datos
    
    ' Definir la celda donde se pegarán los datos en la hoja de destino
    Set celdaDestino = wsDestino.Range("H8")
    
    
  
    
    ' Copiar y pegar valores
    rangoCopiar.Copy
    celdaDestino.PasteSpecial Paste:=xlPasteValues
    
    ' Ajustar columnas en la hoja destino
    wsDestino.Columns("H").AutoFit
    
    
    Set rangoCopiar = wsOrigen.Range("H86:H137") ' Cambia este rango según tus datos
    
    ' Definir la celda donde se pegarán los datos en la hoja de destino
    Set celdaDestino = wsDestino.Range("H82")
    
    ' Copiar y pegar valores
    rangoCopiar.Copy
    celdaDestino.PasteSpecial Paste:=xlPasteValues
    
    ' Ajustar columnas en la hoja destino
    wsDestino.Columns("H").AutoFit
    
    
        ' Establecer la hoja donde se escribirá la hora
    Set ws = ThisWorkbook.Sheets("MAR.DIC.")
    
    ' Escribir la hora en la celda H41
    ws.Range("H50").Value = TimeValue("07:00")
    
      ' Escribir la hora en la celda H51
    ws.Range("H51").Value = TimeValue("11:30")
    
        ws.Range("H66").Value = TimeValue("07:00")
    
      ' Escribir la hora en la celda H51
    ws.Range("H67").Value = TimeValue("12:30")
    
       ' Escribir la hora en la celda H41
    ws.Range("H124").Value = TimeValue("07:00")
    
      ' Escribir la hora en la celda H51
    ws.Range("H125").Value = TimeValue("13:00")
    
        ws.Range("H140").Value = TimeValue("07:00")
    
      ' Escribir la hora en la celda H51
    ws.Range("H141").Value = TimeValue("13:00")
    
    ' Mostrar mensaje
    MsgBox "Los datos se copiaron correctamente de '" & wsOrigen.Name & "' a '" & wsDestino.Name & "'.", vbInformation
End Sub
