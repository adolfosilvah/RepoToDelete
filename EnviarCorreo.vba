Sub EnviarCorreoCierre()
    Dim OutlookApp As Object
    Dim Correo As Object
    Dim Rango As Range
    Dim Cuerpo As String
    Dim Mes As String
    Dim Inspector As Object
    Dim WordDoc As Object
    Dim RangoParaInsertar As Object
    Dim saludos As String
    Dim UltimaFila As Long
    
    ' Definir el mes
    Mes = InputBox("Por favor, ingrese el nombre del mes:", "Solicitud de Mes")
    saludos = SaludoPorHora
    
    ' Crear una nueva instancia de Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    Set Correo = OutlookApp.CreateItem(0)
    
    ' Definir el rango de celdas
    UltimaFila = ThisWorkbook.Sheets("resumen").Cells(ThisWorkbook.Sheets("resumen").Rows.Count, "H").End(xlUp).Row
    Set Rango = ThisWorkbook.Sheets("Resumen").Range("B1:H"&UltimaFila)
    Rango.Copy

    ' Construir el cuerpo del correo
    Cuerpo = saludos
    Cuerpo = Cuerpo & "<p>Adjunto los siguientes documentos:</p>"
    Cuerpo = Cuerpo & "<ul>"
    Cuerpo = Cuerpo & "<li>Inventario total</li>"
    Cuerpo = Cuerpo & "<li>Documento en Excel</li>"
    Cuerpo = Cuerpo & "<li>Resumen de inventario pies/cargas</li>"
    Cuerpo = Cuerpo & "</ul>"
    Cuerpo = Cuerpo & ReporteFOB
    Cuerpo = Cuerpo & "<p>Detalles del mes de " & Mes & ":</p>"
    Cuerpo = Cuerpo & "<p>[AQUÍ VA EL RANGO]</p>" ' Marcador para el rango
    Cuerpo = Cuerpo & "<p>Sin más a que hacer referencia me despido atento a sus comentarios o dudas en cuanto a la información aportada.</p>"
    Cuerpo = Cuerpo & "<p>Saludos.</p>"
    
    ' Configurar el correo
    With Correo
        .To = "jhinojosa@longkeda-int.com; 'Marlen Rojo' <mrojo@sinoenergycorp.com>; mmarron@longkeda-int.com"
        .CC = "'Ernesto Garcia' <Egarcia@sinoenergycorp.com>; zmaroun@sinoenergycorp.com; Habran Perez <hperez@sinoenergycorp.com>; asilva@longkeda-int.com; lsalazar@longkeda-int.com; orodriguez@longkeda-int.com; lperdomo@sinoenergycorp.com"
        .Subject = "Cierre de mes " & Mes & " 2024"
        .HTMLBody = Cuerpo
        .Display ' Mostrar el correo para editar
    End With
    
    ' Insertar el rango copiado en el cuerpo del correo
    Set Inspector = Correo.GetInspector
    Set WordDoc = Inspector.WordEditor
    Set RangoParaInsertar = WordDoc.Content
    RangoParaInsertar.Find.Execute FindText:="[AQUÍ VA EL RANGO]"
    RangoParaInsertar.Paste ' Pegar el rango desde el portapapeles

    ' Limpiar objetos
    Set Correo = Nothing
    Set OutlookApp = Nothing
End Sub

Function ReporteFOB() As String
    Dim UltimaFila As Long
    Dim ValorFOB As Double

    ' Encuentra la última fila con datos en la columna H
    UltimaFila = ThisWorkbook.Sheets("resumen").Cells(ThisWorkbook.Sheets("resumen").Rows.Count, "H").End(xlUp).Row

    ' Obtiene el valor de la última celda con datos en la columna H
    ValorFOB = ThisWorkbook.Sheets("resumen").Cells(UltimaFila, "H").Value

If ValorFOB > 0 Then
    ReporteFOB = "Para el mes de acuerdo con el cierre administrativo en cuanto a los consumibles de la línea de WL (FOB) $" & ValorFOB
 Else
    ReporteFOB = "<p>Para el mes no se generó gastos de acuerdo con el cierre administrativo en cuanto a los consumibles de la línea de WL (FOB) $ 0 (sin actividades)"
End If

End Function




Function SaludoPorHora() As String
    Dim HoraActual As Double
    Dim Saludo As String
    
    ' Obtiene la hora actual
    HoraActual = Time
    
    ' Determina el saludo basado en la hora actual
    If Hour(HoraActual) >= 0 And Hour(HoraActual) < 12 Then
        Saludo = "Buen día Sres."
    ElseIf Hour(HoraActual) >= 12 And Hour(HoraActual) < 18 Then
        Saludo = "Buenas tardes Sres."
    Else
        Saludo = "Buenas noches Sres."
    End If
    
    ' Devuelve el saludo
    SaludoPorHora = Saludo
End Function





