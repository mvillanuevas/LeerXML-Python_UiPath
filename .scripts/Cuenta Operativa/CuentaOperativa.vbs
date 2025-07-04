'On Error Resume Next

'Set objArgs = WScript.Arguments

'WorkbookPathRexmex = objArgs(0)
'WorkbookSheetRexmex = objArgs(1)
'WorkbookPathN = objArgs(2)
'WorkbookSheetNomina = objArgs(3)
'anio = objArgs(4)
'mes = objArgs(5)

'anio = CInt(anio)
'mes = CInt(mes)

WorkbookPathRexmex = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\REXMEX - Cuenta Operativa 2025_120525.xlsx"
WorkbookSheetRexmex = "Cuenta Operativa"
WorkbookPathN = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\Nomina Abr25_conNoAcreedor.xlsx"
WorkbookSheetNomina = "NOMINA"
anio = 2025
mes = 3

'Genera un objeto de tipo Excel Application
Set objExcel = CreateObject("Excel.Application")

'Parámetro para indicar si se quiere visible la aplicación de Excel
objExcel.Application.Visible = True
'Evita movimiento de pantalla
objExcel.Application.ScreenUpdating = True
'Parámetro evitar mostrar pop ups de Excel
objExcel.Application.DisplayAlerts = False
' Desactiva eventos de Excel
objExcel.Application.EnableEvents = False

'Abre libro Excel
Set objWorkbookPathNomina = objExcel.Workbooks.Open(WorkbookPathN)
Set objWorkbookSheetNomina = objWorkbookPathNomina.Worksheets(WorkbookSheetNomina)

Set objWorkbookPathRexmex = objExcel.Workbooks.Open(WorkbookPathRexmex)
Set objWorkbookSheetRexmex = objWorkbookPathRexmex.Worksheets(WorkbookSheetRexmex)

' Verificar si los filtros están activos en la fila 1, si no, activarlos
If objWorkbookSheetNomina.AutoFilterMode Then
    objWorkbookSheetNomina.AutoFilterMode = False
End If

If objWorkbookSheetRexmex.AutoFilterMode Then
    objWorkbookSheetRexmex.AutoFilterMode = False
End If

Dim ultimoDiaMes
ultimoDiaMes = DateSerial(anio, mes + 1, 0)

dim filesys
Set filesys = CreateObject("Scripting.FileSystemObject")
nombreArchivo = filesys.GetFileName(WorkbookPathN)

nombreArchivo = Replace(nombreArchivo, ".XLSX", "")
nombreArchivo = Replace(nombreArchivo, ".xlsx", "")

' Encontrar la Ultima fila con datos en la columna a filtrar
lastRow = objWorkbookSheetNomina.Cells(objWorkbookSheetNomina.Rows.Count, 2).End(-4162).Row + 1 ' -4162 = xlUp
lastCol = objWorkbookSheetNomina.Cells(3, objWorkbookSheetNomina.Columns.Count).End(-4159).Column ' -4159 = xlToLeft

' Iterar la columna A hasta encontrar una celda vacia
For i = 10 To lastRow
	If Trim(objWorkbookSheetNomina.Cells(i, 1).Value) = "" Then
        firstRow = i + 1
		Exit For
    End If
Next

Set uniqueDict = CreateObject("Scripting.Dictionary")
Set pepDict = CreateObject("Scripting.Dictionary")

For i = firstRow to lastRow
	' Obtener la Ultima fila con datos en la hoja de REXMEX
    lastRowR = objWorkbookSheetRexmex.Cells(objWorkbookSheetRexmex.Rows.Count, 1).End(-4162).Row ' -4162 = xlUp
	
	' Escribe fomula R1C1
	flag = SetFormulaR1C1(objWorkbookSheetRexmex, lastRowR)
	
	If objWorkbookSheetNomina.Cells(i, 2).Value = "" And objWorkbookSheetNomina.Cells(i, 3).Value = "" And objWorkbookSheetNomina.Cells(i, 5).Value = "" Then
		For j = 9 to lastCol
			' Verifica si la celda no est� vacia y no es cero
			If objWorkbookSheetNomina.Cells(i, j).Value <> "" And objWorkbookSheetNomina.Cells(i, j).Value <> 0 Then
				' Escribe el valor de la celda en la hoja de REXMEX
				pep = objWorkbookSheetNomina.Cells(3, j).Value
				valorPEP = objWorkbookSheetNomina.Cells(i, j).Value
				rowR = j - 8
				objWorkbookSheetRexmex.Cells(lastRowR + rowR, 1).Value = "PEP"
				objWorkbookSheetRexmex.Cells(lastRowR + rowR, 2).Value = "FP"
				objWorkbookSheetRexmex.Cells(lastRowR + rowR, 3).Value = "MX29"
				objWorkbookSheetRexmex.Cells(lastRowR + rowR, 4).Value = anio
				objWorkbookSheetRexmex.Cells(lastRowR + rowR, 5).Value = mes
				objWorkbookSheetRexmex.Cells(lastRowR + rowR, 6).Value = ultimoDiaMes
				objWorkbookSheetRexmex.Cells(lastRowR + rowR, 7).Value = pep
				objWorkbookSheetRexmex.Cells(lastRowR + rowR, 16).Value = nombreArchivo
				objWorkbookSheetRexmex.Cells(lastRowR + rowR, 20).Value = valorAcreedor
				objWorkbookSheetRexmex.Cells(lastRowR + rowR, 36).Value = valorPEP
				objWorkbookSheetRexmex.Cells(lastRowR + rowR, 54).Value = totalFactura
				objWorkbookSheetRexmex.Cells(lastRowR + rowR, 56).Value = uuid
				objWorkbookSheetRexmex.Cells(lastRowR + rowR, 57).Value = fecha
				objWorkbookSheetRexmex.Cells(lastRowR + rowR, 60).Value = numDocto
			End If
		Next
		' Autofill de la columna 24 a 31
		objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(lastRowR + 1, 24), objWorkbookSheetRexmex.Cells(lastRowR + 1, 31)).AutoFill _
    		objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(lastRowR + 1, 24), objWorkbookSheetRexmex.Cells(lastRowR + rowR, 31))
		' Autofill de la columna 33 a 35
		objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(lastRowR + 1, 33), objWorkbookSheetRexmex.Cells(lastRowR + 1, 35)).AutoFill _
			objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(lastRowR + 1, 33), objWorkbookSheetRexmex.Cells(lastRowR + rowR, 35))
		' Autofill de la columna 37 a 51
		objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(lastRowR + 1, 37), objWorkbookSheetRexmex.Cells(lastRowR + 1, 51)).AutoFill _
			objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(lastRowR + 1, 37), objWorkbookSheetRexmex.Cells(lastRowR + rowR, 51))
		' Autofill de la columna 64 a 66
		objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(lastRowR + 1, 64), objWorkbookSheetRexmex.Cells(lastRowR + 1, 66)).AutoFill _
			objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(lastRowR + 1, 64), objWorkbookSheetRexmex.Cells(lastRowR + rowR, 66))
		' Autofill de la columna 68 a 70
		objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(lastRowR + 1, 68), objWorkbookSheetRexmex.Cells(lastRowR + 1, 70)).AutoFill _
			objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(lastRowR + 1, 68), objWorkbookSheetRexmex.Cells(lastRowR + rowR, 70))

		' Itera sobre la columna 1 y elimina las filas con valores vacios en la columna 1
		For j = lastRowR + rowR To lastRowR Step -1
			If objWorkbookSheetRexmex.Cells(j, 1).Value = "" Then
				objWorkbookSheetRexmex.Rows(j).Delete
			End If
		Next

	ElseIf objWorkbookSheetNomina.Cells(i, 5).Value <> "" And objWorkbookSheetNomina.Cells(i, 7).Value <> "" Then
		totalFactura = objWorkbookSheetNomina.Cells(i, 5).Value
		valorAcreedor = objWorkbookSheetNomina.Cells(i, 1).Value
		numDocto = objWorkbookSheetNomina.Cells(i, 7).Value
	ElseIf objWorkbookSheetNomina.Cells(i, 3).Value <> "" And objWorkbookSheetNomina.Cells(i, 4).Value <> "" Then
		uuid = objWorkbookSheetNomina.Cells(i, 3).Value
		fecha = objWorkbookSheetNomina.Cells(i, 4).Value
	End If
Next


' Guardar y cerrar los libros de trabajo
objWorkbookPathRexmex.Save
objWorkbookPathRexmex.Close

objWorkbookPathNomina.Save
objWorkbookPathNomina.Close
' Cerrar la aplicación de Excel
objExcel.Quit

'Devuelve el error en caso de
If Err.Number <> 0 Then
    Msg = "Error was generated by " & Err.Source & ". " & Err.Description
    WScript.StdOut.WriteLine Msg
End if

Function SetFormulaR1C1(objWorkbookSheetRexmex, lastRowR)
	
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 24).FormulaR1C1 = "=VLOOKUP(RC[-21],WI!R9C2:R14C3,2,0)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 26).FormulaR1C1 = "=+RC[-19]"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 28).FormulaR1C1 = "=VLOOKUP(RC[-2],'Elementos PEP'!C3:C7,5,FALSE)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 29).FormulaR1C1 = "=VLOOKUP(RC[-3],'Elementos PEP'!C3:C7,2,FALSE)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 30).FormulaR1C1 = "=VLOOKUP(RC[-4],'Elementos PEP'!C3:C7,3,FALSE)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 31).FormulaR1C1 = "=VLOOKUP(RC[-5],'Elementos PEP'!C3:C7,4,FALSE)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 33).FormulaR1C1 = "=+RC[-13]"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 34).FormulaR1C1 = "=VLOOKUP(RC[-1],Proveedores!C[-32]:C[-30],2,0)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 35).FormulaR1C1 = "=VLOOKUP(RC[-2],Proveedores!C[-33]:C[-31],3,0)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 37).FormulaR1C1 = "=+RC[-23]"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 38).FormulaR1C1 = "=IFERROR(VLOOKUP(RC24,WI!R9C3:R14C9,2,0)*RC36,0)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 39).FormulaR1C1 = "=IFERROR(VLOOKUP(RC24,WI!R9C3:R14C9,2,0)*RC51,0)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 40).FormulaR1C1 = "=IFERROR(VLOOKUP(RC24,WI!R9C3:R14C9,3,0)*RC36,0)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 41).FormulaR1C1 = "=IFERROR(VLOOKUP(RC24,WI!R9C3:R14C9,3,0)*RC51,0)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 42).FormulaR1C1 = "=IFERROR(VLOOKUP(RC24,WI!R9C3:R14C9,4,0)*RC36,0)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 43).FormulaR1C1 = "=IFERROR(VLOOKUP(RC24,WI!R9C3:R14C9,4,0)*RC51,0)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 44).FormulaR1C1 = "=IFERROR(VLOOKUP(RC24,WI!R9C3:R14C9,5,0)*RC36,0)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 45).FormulaR1C1 = "=IFERROR(VLOOKUP(RC24,WI!R9C3:R14C9,5,0)*RC51,0)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 46).FormulaR1C1 = "=IFERROR(VLOOKUP(RC24,WI!R9C3:R14C9,6,0)*RC36,0)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 47).FormulaR1C1 = "=IFERROR(VLOOKUP(RC24,WI!R9C3:R14C9,6,0)*RC51,0)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 48).FormulaR1C1 = "=IFERROR(VLOOKUP(RC24,WI!R9C3:R14C9,7,0)*RC36,0)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 49).FormulaR1C1 = "=IFERROR(VLOOKUP(RC24,WI!R9C3:R14C9,7,0)*RC51,0)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 50).FormulaR1C1 = "=+RC[-14]"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 51).FormulaR1C1 = "=+RC[-1]*0.16"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 64).FormulaR1C1 = "=+RC[-55]"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 65).FormulaR1C1 = "=VLOOKUP(RC[-39],'Elementos PEP'!C[-62]:C[-57],6,FALSE)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 66).FormulaR1C1 = "=VLOOKUP(RC[-32],Proveedores!C[-63]:C[-61],3,0)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 68).FormulaR1C1 = "=IF(RC[-31]=" & Chr(34) & "MXN" & Chr(34) & ",(RC[-18]+RC[-17])/VLOOKUP('Cuenta Operativa'!RC[-41],TC!C[-67]:C[-62],4,0),IF(RC[-31]=" & Chr(34) & "EUR" & Chr(34) & ",('Cuenta Operativa'!RC[-18]+'Cuenta Operativa'!RC[-17])*VLOOKUP('Cuenta Operativa'!RC[-41],TC!C[-67]:C[-62],5,0),'Cuenta Operativa'!RC[-18]+'Cuenta Operativa'!RC[-17]))"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 69).FormulaR1C1 = "=IFERROR(((RC[-8]>1)*IF(RC[-32]=" & Chr(34) & "MXN" & Chr(34) & ",(RC[-19]+RC[-18])/VLOOKUP('Cuenta Operativa'!RC[-8],TC!C[-68]:C[-63],4,0),IF(RC[-32]=" & Chr(34) & "EUR" & Chr(34) & ",('Cuenta Operativa'!RC[-19]+'Cuenta Operativa'!RC[-18])*VLOOKUP('Cuenta Operativa'!RC[-8],TC!C[-68]:C[-63],5,0),'Cuenta Operativa'!RC[-19]+'Cuenta Operativa'!RC[-18]))),0)"
	objWorkbookSheetRexmex.Cells(lastRowR + 1, 70).FormulaR1C1 = "=IFERROR(((RC[-9]>1)*IF(RC[-33]=" & Chr(34) & "MXN" & Chr(34) & ",RC[-17]/VLOOKUP('Cuenta Operativa'!RC[-9],TC!C[-69]:C[-64],4,0),IF(RC[-33]=" & Chr(34) & "EUR" & Chr(34) & ",RC[-17]*VLOOKUP('Cuenta Operativa'!RC[-9],TC!C[-69]:C[-64],5,0),RC[-17]))),0)"

	SetFormulaR1C1 = True
End Function