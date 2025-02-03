Attribute VB_Name = "M�dulo1"
Sub ModificarXLCerrados()
'Declaramos variables
Dim Archivo As Application
Dim Celda As Object
Dim NombreArchivo As String
'
'Creamos el objecto Excel
Set Archivo = CreateObject("Excel.Application")
'
With Archivo
    '
    'Recorremos cada celda de la selecci�n para tomar el nombre de cada archivo
    For Each Celda In Selection
        NombreArchivo = Celda.Value
        '
        'Validamos si el archivo ya est� abierto
       
            '
        With .Workbooks.Open(NombreArchivo)
            'Hacemos las modificaciones en el archivo
            .Worksheets("Hoja1").Range("A1").Value = "Total"
            .Worksheets("Hoja1").Range("A2").Value = 3000
            'Cerramos el archivo guardando cambios
            .Close SaveChanges:=True
        End With
        
        '
    Next Celda
    '
    'Cerramos la aplicaci�n de Excel
    .Quit
End With
End Sub


