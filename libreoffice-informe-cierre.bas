REM  *****  BASIC  *****

Sub Main
	dim doc as Object 
	dim sheet as Object
	
	filas 	   = 18 ' cantidad filas antes del footer
	filaFooter = filas + 1
	
	doc 		= ThisComponent
	sheet 		= doc.sheets(0)
	columnas 	= sheet.getColumns()
	
	colorPrimario = RGB(7, 25, 82)
	colorTextPrimario = RGB(255, 255, 255)
	
		
	range = sheet.getCellRangeByName("A1:E1")
	range.cellBackColor = colorPrimario ' color de fondo
	range.charColor 	= colorTextPrimario ' color de texto
	range.charWeight 	= com.sun.star.awt.FontWeight.BOLD 
	
	' tamaño de columnas
	columnas.getByIndex(0).width = 500
	columnas.getByIndex(1).width = 15000
	columnas.getByIndex(2).width = 900
	columnas.getByIndex(3).width = 900
	columnas.getByIndex(4).width = 15000
	
	' Ajuste de texto
	columnas.getByIndex(1).isTextWrapped = true
	columnas.getByIndex(4).isTextWrapped = true
	
	' Alineación
	rangoFooter = sheet.getCellRangeByName("A" & filaFooter & ":E" & filaFooter)
	rangoFooter.cellBackColor 	= colorPrimario
	rangoFooter.charColor 		= colorTextPrimario
	rangoFooter.charWeight 		= com.sun.star.awt.FontWeight.BOLD 
	sheet.getCellRangeByName("A" & filaFooter & ":C" & filaFooter).horiJustify = com.sun.star.table.CellHoriJustify.RIGHT
	
	' Reemplazo de texto	
	for i = 2 to filas
		celda = sheet.getCellByPosition(4, i)
		texto = celda.getString()
		textoMod = replace(texto, "- ", Chr(13)&"- ")
		celda.setString(textoMod)
	next i
End Sub

Sub Prueba
	n = NotaFinal(7, 7, 7, 0, 7, 7, 7)
	MsgBox "La Nota es " & n
End Sub

Function NotaFinal(p1, p2, p3, f1, f2, f3, examen) As Double
	' Pruebas
	nota = nota + p1 * 0.15
	nota = nota + p2 * 0.15
	nota = nota + p3 * 0.15
	
	' Foros			
	nota = nota + IIf(f1 = 0, 1, f1) * 0.05
	nota = nota + IIf(f2 = 0, 1, f2) * 0.05
	nota = nota + IIf(f3 = 0, 1, f3) * 0.05
	
	nota = nota + examen * 0.4
	
	NotaFinal = nota
End Function

Function NotaChilena(puntajeObtenido As Double) As Double
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oCell As Object    
    Dim puntajeMax As Double
    Dim notaMax As Double
    Dim notaMin As Double
    Dim notaAprobatoria As Double
    Dim exigencia As Double
    Dim notaCalculada As Double
    
    ' Parámetros de la escala de notas
    notaMax = 7
    notaMin = 1
    notaAprobatoria = 4
    exigencia = 0.6
    puntajeMax = 100
    'puntajeObtenido = 30
    
    ' Obtener referencia al documento actual
    oDoc = ThisComponent
    
    ' Obtener referencia a la hoja activa
    oSheet = oDoc.getCurrentController.getActiveSheet()
    
    ' Asumimos que el puntaje obtenido está en la celda A1 y el puntaje máximo en la celda A2
    'puntajeObtenido = oSheet.getCellByPosition(0, 0).Value
        
    ' Calcular la nota
    If puntajeObtenido >= 60 Then
        notaCalculada = 3 * ((puntajeObtenido-60)/40) + 4
    Else
        notaCalculada = 3 * ((puntajeObtenido)/60) + 1
    End If
    
    'MsgBox "Nota calculada: " & notaCalculada
    NotaChilena = notaCalculada
End Function

Sub GenerarInformes
	filaInicioPegar = 23
	Informe("INFORME P1", 14, filaInicioPegar)
	Informe("INFORME P2", 15, filaInicioPegar)
	Informe("INFORME P3", 16, filaInicioPegar)
	Informe("INFORME EXAMEN y EX REP", 17, filaInicioPegar)
	Informe("BBDD NOTA FINAL", 6, 3)
	ActaFinal()
End Sub

Sub ActaFinal
	oSheetData = ThisComponent.Sheets.getByName("Data")
	oCursor = oSheetData.createCursorByRange(oSheetData.getCellRangeByName("B2"))
    oCursor.gotoEndOfUsedArea(False)    
    lastRow = oCursor.RangeAddress.EndRow
    firstRow = 1 ' 0 indexed --> 1 == 2
    rows = lastRow - firstRow + 1
    
	Dim notasFinales(0 To rows-1, 0 To 0) as Double 
	Dim notasExamenes(0 To rows-1, 0 To 0) as Double 
	Dim notas1(0 To rows-1, 0 To 0) as Double 
	Dim notas2(0 To rows-1, 0 To 0) as Double 
	Dim notas3(0 To rows-1, 0 To 0) as Double 
	
	Dim nombres(0 To rows-1, 0 To 0) as String
	Dim ruts(0 To rows-1, 0 To 0) as String
	
	For i = 0 To rows-1		
		colP1 = 14
		colP2 = 15
		colP3 = 16
		colEx = 17
		colF1 = 8
		colF2 = 10
		colF3 = 12
		p1 = NotaChilena( oSheetData.getCellByPosition(colP1, i+1 ).getValue() )
		p2 = NotaChilena( oSheetData.getCellByPosition(colP2, i+1 ).getValue() )
		p3 = NotaChilena( oSheetData.getCellByPosition(colP3, i+1 ).getValue() )
		examen = NotaChilena( oSheetData.getCellByPosition(colEx, i+1 ).getValue() )
		f1 = oSheetData.getCellByPosition(colF1, i+1 ).getValue()
		f2 = oSheetData.getCellByPosition(colF2, i+1 ).getValue()
		f3 = oSheetData.getCellByPosition(colF3, i+1 ).getValue()
		notasFinales(i, 0) = NotaFinal(p1, p2, p3, f1, f2, f3, examen)
		notasExamenes(i, 0) = examen
		notas1(i, 0) = p1
		notas2(i, 0) = p2
		notas3(i, 0) = p3
		' ==========================================================
		nombre 		= oSheetData.getCellByPosition(1, i+1 ).getString() ' col B = 1
		apellido 	= oSheetData.getCellByPosition(0, i+1 ).getString() ' col A = 0
		nombreCompleto = nombre & " " & apellido
		nombres(i, 0) = nombreCompleto
		' ==========================================================		
		ruts(i, 0) = oSheetData.getCellByPosition(3, i+1).getString()
	Next i
	
	acta = ThisComponent.Sheets.getByName("ACTA FINAL")
	actaFirstRow = 8
	actaLastRow = actaFirstRow + rows - 1
	
    colB = 1 ' nombreCompleto
    oCellRangeDest = acta.getCellRangeByPosition(colB, actaFirstRow, colB, actaLastRow)
    oCellRangeDest.setDataArray( nombres )  

	colA = 0 ' RUT
    oCellRangeDest = acta.getCellRangeByPosition(colA, actaFirstRow, colA, actaLastRow)
    oCellRangeDest.setDataArray( ruts )
    
    colC = 2 ' p1
    oCellRangeDest = acta.getCellRangeByPosition(colC, actaFirstRow, colC, actaLastRow)
    oCellRangeDest.setDataArray( notas1 )
    
    colD = 3 ' p2
    oCellRangeDest = acta.getCellRangeByPosition(colD, actaFirstRow, colD, actaLastRow)
    oCellRangeDest.setDataArray( notas2 )
    
    colE = 4 ' p3
    oCellRangeDest = acta.getCellRangeByPosition(colE, actaFirstRow, colE, actaLastRow)
    oCellRangeDest.setDataArray( notas3 )
    
    colF = 5 ' examen
    oCellRangeDest = acta.getCellRangeByPosition(colF, actaFirstRow, colF, actaLastRow)
    oCellRangeDest.setDataArray( notasExamenes )
    
    colG = 6 ' notaFinal
	oCellRangeDest = acta.getCellRangeByPosition(colG, actaFirstRow, colG, actaLastRow)
    oCellRangeDest.setDataArray( notasFinales )
End Sub

Sub Informe(nombreHojaInforme as String, colNota as Long, filaInicioPegar as Long)
	Dim oDoc As Object
    Dim oSheetData As Object
    Dim oSheetInformeP1 As Object
    Dim oSheetInformeP2 As Object
    Dim oSheetInformeP3 As Object
    Dim oCellRangeSource As Object
    Dim oCellRangeDest As Object
    Dim oCursor As Object
    Dim lastRow As Long
    
    ' Referencia al documento actual
    oDoc = ThisComponent
    
    ' Referencia a la hoja "Data"
    oSheetData = oDoc.Sheets.getByName("Data")
    
    ' Determinar la última fila con datos en la columna B
    oCursor = oSheetData.createCursorByRange(oSheetData.getCellRangeByName("B2"))
    oCursor.gotoEndOfUsedArea(False)
    lastRow = oCursor.RangeAddress.EndRow
    
    ' Referencia a la hoja "INFORME P1"
    oSheetInformeP1 = oDoc.Sheets.getByName( nombreHojaInforme )    
    informeLastRow = filaInicioPegar + lastRow - 1
    ' ===================================================================
    ' Copiar desde B2 hasta la última fila en la hoja "Data"
    oCellRangeSource = oSheetData.getCellRangeByPosition(1, 1, 1, lastRow) ' (columna B, fila 2) -> (columna B, última fila)
    
    ' Pegar en A24 en adelante en la hoja "INFORME P1"
    oCellRangeDest = oSheetInformeP1.getCellRangeByPosition(0, filaInicioPegar, 0, informeLastRow) ' (columna A, fila 24) -> (columna A, fila 24 + lastRow - 1)
    oCellRangeDest.setDataArray(oCellRangeSource.getDataArray())
    
    ' ===================================================================
    ' Copiar desde A2 hasta la última fila en la hoja "Data"
    oCellRangeSource = oSheetData.getCellRangeByPosition(0, 1, 0, lastRow) ' (columna A, fila 2) -> (columna A, última fila)
    
    ' Pegar en B24 en adelante en la hoja "INFORME P1"
    oCellRangeDest = oSheetInformeP1.getCellRangeByPosition(1, filaInicioPegar, 1, informeLastRow) ' (columna B, fila 24) -> (columna B, fila 24 + lastRow - 1)
    oCellRangeDest.setDataArray(oCellRangeSource.getDataArray())

	' ===================================================================
	' Copiar desde C2 hasta la última fila en la hoja "Data"
    oCellRangeSource = oSheetData.getCellRangeByPosition(2, 1, 2, lastRow) ' (columna C, fila 2) -> (columna C, última fila)
    
    ' Pegar en C24 en adelante en la hoja "INFORME P1"
    oCellRangeDest = oSheetInformeP1.getCellRangeByPosition(2, filaInicioPegar, 2, informeLastRow) ' (columna C, fila 24) -> (columna C, fila 24 + lastRow - 1)
    oCellRangeDest.setDataArray(oCellRangeSource.getDataArray())
    
    ' Pegar en G24 en adelante en la hoja "INFORME P1"
    oCellRangeDest = oSheetInformeP1.getCellRangeByPosition(6, filaInicioPegar, 6, informeLastRow) ' (columna G, fila 24) -> (columna G, fila 24 + lastRow - 1)
    oCellRangeDest.setDataArray(oCellRangeSource.getDataArray())
    
    ' ===================================================================
	' Copiar desde D2 hasta la última fila en la hoja "Data"
    oCellRangeSource = oSheetData.getCellRangeByPosition(3, 1, 3, lastRow) ' (columna D, fila 2) -> (columna D, última fila)
    
    ' Pegar en D24 en adelante en la hoja "INFORME P1"
    oCellRangeDest = oSheetInformeP1.getCellRangeByPosition(3, filaInicioPegar, 3, informeLastRow) ' (columna D, fila 24) -> (columna D, fila 24 + lastRow - 1)
    oCellRangeDest.setDataArray(oCellRangeSource.getDataArray())
    ' ===================================================================
	' Copiar desde O2 hasta la última fila en la hoja "Data"	
	' colNota=14 
    oCellRangeSource = oSheetData.getCellRangeByPosition(colNota, 1, colNota, lastRow) ' (columna O, fila 2) -> (columna O, última fila)
        
    ' Pegar en I24 en adelante en la hoja "INFORME P1"
    colI=8
    oCellRangeDest = oSheetInformeP1.getCellRangeByPosition(colI, filaInicioPegar, colI, informeLastRow) ' (columna D, fila 24) -> (columna D, fila 24 + lastRow - 1)
    oCellRangeDest.setDataArray(oCellRangeSource.getDataArray())
    ' ===================================================================
    For i = (filaInicioPegar+1) To informeLastRow + 1
    	puntajeObtenido = oSheetInformeP1.getCellByPosition(colI, i-1 ).getValue()
    	nota = NotaChilena( puntajeObtenido )
    	oSheetInformeP1.getCellByPosition(colI-1, i-1 ).setValue( nota )
    Next i
    ' ===================================================================
    ' MsgBox "Datos copiados con éxito."
End Sub
