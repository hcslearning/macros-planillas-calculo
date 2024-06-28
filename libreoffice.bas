REM  *****  BASIC  *****

Sub Main
	dim doc as Object 
	dim sheet as Object
	
	filas 	   = 12 ' cantidad filas antes del footer
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

