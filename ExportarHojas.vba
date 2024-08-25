Sub ExportarHojas()
    Dim wbNuevo As Workbook
    Dim ruta As String
    Dim nombreArchivo As String
    Dim hojas As Variant
    Dim i As Integer
    Dim carpetaDestino As String
    Dim hoja As Worksheet
    Dim pdfRuta As String
    Dim Mes As String
    
    Application.ScreenUpdating = False

    ' Definir las hojas a exportar
    hojas = Array("Resumen Pies x Cargas", "Resumen", "Detalles de Consumo", "Consumo Operacional", "Disponibilidad")
    
    ' Obtener la ruta del libro activo
    ruta = ThisWorkbook.Path
    

    
    ' Pedir el nombre del nuevo archivo
     Mes = InputBox("Introduce el nombre del mes del cierre :")
    nombreArchivo = "Cierre de mes " & Mes & " 2024"
    ' Crear un nuevo libro
    Set wbNuevo = Workbooks.Add
    
       ' Definir la carpeta de destino
    carpetaDestino = ruta & "\Cierres de mes"\Mes 
    ' Copiar las hojas al nuevo libro
    For i = LBound(hojas) To UBound(hojas)
        ThisWorkbook.Sheets(hojas(i)).Copy After:=wbNuevo.Sheets(wbNuevo.Sheets.Count)
    Next i
    
    ' Eliminar la hoja en blanco que se crea por defecto
    Application.DisplayAlerts = False
    wbNuevo.Sheets(1).Delete
    Application.DisplayAlerts = True
    
    ' Guardar el nuevo libro en la carpeta "Cierres de mes"
    wbNuevo.SaveAs fileName:=carpetaDestino & "\" & nombreArchivo & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    
    ' Exportar las hojas "Resumen Pies x Cargas" y "Detalles de Consumo" a PDF
    For Each hoja In ThisWorkbook.Sheets(Array("Resumen Pies x Cargas", "Disponibilidad"))
        pdfRuta = carpetaDestino & "\" & hoja.Name & " " & Mes & " 2024" & ".pdf"
        hoja.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfRuta, Quality:=xlQualityStandard
    Next hoja
    
    ' Cerrar el nuevo libro
    wbNuevo.Close
    
    Application.ScreenUpdating = True
    
    MsgBox "Las hojas han sido exportadas y guardadas en " & carpetaDestino & "\" & nombreArchivo & ".xlsx" & vbCrLf & _
           "Los PDFs han sido guardados en la misma ubicación."
End Sub

