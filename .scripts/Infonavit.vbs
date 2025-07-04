On Error Resume Next

Set objArgs = WScript.Arguments

WorkbookPathIMSS = objArgs(0)
WorkbookSheetIMSS = objArgs(1)
WorkbookPathNomina = objArgs(2)
WorkbookSheetNomina = objArgs(3)
WorkbookPathCatalogo = objArgs(4)
WorkbookSheetCatalogo = objArgs(5)

'WorkbookPathIMSS = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\03 IMSS.xlsx"
'WorkbookSheetIMSS = "Sheet1"
'WorkbookPathNomina = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\Nomina abr25.XLSX"
'WorkbookSheetNomina = "NOMINA"
'WorkbookPathCatalogo = "C:\Users\HE678HU\OneDrive - EY\Documents\UiPath\Leer_Facturas_Nomina\.templates\Catalogo Bloques.xlsx"
'WorkbookSheetCatalogo = "Centro"

'Genera un objeto de tipo Excel Application
Set objExcel = CreateObject("Excel.Application")

'Parámetro para indicar si se quiere visible la aplicación de Excel
objExcel.Application.Visible = False
'Evita movimiento de pantalla
objExcel.Application.ScreenUpdating = False
'Parámetro evitar mostrar pop ups de Excel
objExcel.Application.DisplayAlerts = False

'Abre libro Excel
Set objWorkbookIMSS = objExcel.Workbooks.Open(WorkbookPathIMSS)
Set objWorksheetIMSS = objWorkbookIMSS.Worksheets(WorkbookSheetIMSS)

' Verificar si los filtros están activos en la fila 1, si no, activarlos
If Not objWorksheetIMSS.AutoFilterMode Then
    objWorksheetIMSS.Rows(1).AutoFilter
End If

colToFilter = 1 ' Número de columna a filtrar (1 = columna A)
' Encontrar la última fila con datos en la columna a filtrar
lastRow = objWorksheetIMSS.Cells(objWorksheetIMSS.Rows.Count, colToFilter).End(-4162).Row ' -4162 = xlUp

' Aplicar autofiltro con los valores "aa", "bb" y "cc"
objWorksheetIMSS.Range(objWorksheetIMSS.Cells(1, colToFilter), objWorksheetIMSS.Cells(lastRow, colToFilter)).AutoFilter _
                        23, Array("Infonavit"), 7 ' 7 = xlFilterValues

' Copiar las celdas visibles de las columnas X (24), Y (25), Z (26), AA (27)
Dim copyRange
Set copyRange = objWorksheetIMSS.Range("X1:AA" & lastRow).SpecialCells(12) ' 12 = xlCellTypeVisible
copyRange.Copy

' Quitar los filtros aplicados si existen
If objWorksheetIMSS.AutoFilterMode Then
    objWorksheetIMSS.AutoFilterMode = False
End If

' Encontrar la última fila ocupada en la columna AK
Dim lastAK
lastAK = objWorksheetIMSS.Cells(objWorksheetIMSS.Rows.Count, 37).End(-4162).Row + 3

' Pegar en la columna AK (columna 37), desde la fila 1
objWorksheetIMSS.Range("AK" & lastAK).PasteSpecial -4163 ' -4163 = xlPasteAll

' Quitar el modo de corte/copia
objExcel.CutCopyMode = False

' Copiar valores únicos de un rango (por ejemplo, X1:X{lastRow}) a un array
Dim uniqueDict, cell, uniqueList, idx
Set uniqueDict = CreateObject("Scripting.Dictionary")

Dim rngX
Set rngX = objWorksheetIMSS.Range("X2:X" & lastRow)

' Recorrer el rango y agregar valores únicos al diccionario
' (se asume que la primera fila es un encabezado y se comienza desde la segunda fila)
For Each cell In rngX
    If Not IsEmpty(cell.Value) And Trim(cell.Value & "") <> "" Then
        If Not uniqueDict.Exists(cell.Value) Then
            uniqueDict.Add cell.Value, True
        End If
    End If
Next

' Encontrar la última fila ocupada en la columna AK
Dim lastAK2, lastAK3, lastAK4
lastAK2 = objWorksheetIMSS.Cells(objWorksheetIMSS.Rows.Count, 37).End(-4162).Row + 3
lastAK4 = lastAK2
lastAK3 = objWorksheetIMSS.Cells(objWorksheetIMSS.Rows.Count, 37).End(-4162).Row

' Pegar los valores únicos en otra columna
For Each key In uniqueDict.Keys
    objWorksheetIMSS.Cells(lastAK2, 37).Value = key ' 37 = columna AK
    objWorksheetIMSS.Cells(lastAK2, 38).Formula = "=-SUMIF($AK$" & lastAK & ":$AK$" & lastAK3 & ",AK" & lastAK2 & ",$AN$" & lastAK & ":$AN$" & lastAK3 & ")"
    lastAK2 = lastAK2 + 1
Next

'Abre libro Excel
Set objWorkbookN = objExcel.Workbooks.Open(WorkbookPathNomina)
Set objWorksheetN = objWorkbookN.Worksheets(WorkbookSheetNomina)

Set objWorkbookCatalogo = objExcel.Workbooks.Open(WorkbookPathCatalogo)
Set objWorksheetCatalogo = objWorkbookCatalogo.Worksheets(WorkbookSheetCatalogo)

ultimaFilaC = objWorksheetCatalogo.Cells(objWorksheetCatalogo.Rows.Count,1).End(-4162).Row

Const xlPart = 2
Const xlValues = -4163

Set uRange = objWorksheetIMSS.Cells
' Bsuca el valor Linea de Captura INFONAVIT y toma el valor de la siguiente celda
Dim LineaUUID : Set LineaUUID = uRange.Find("nea de Captura INFONAVIT",,xlValues,xlPart)

If Not LineaUUID Is Nothing Then
	CapturaUUID = objWorksheetIMSS.Cells(LineaUUID.Row, LineaUUID.Column + 1).Value
End If

' Establece rango de busqueda en la columna C
Set nRange = objWorksheetN.Range("C:C")
' Busca la celda que contiene "IMSS" en la columna C
Dim IMSS : Set IMSS = nRange.Find("INFONAVIT",,xlValues,xlPart)
Dim iCC, CC

' Si se encuentra la celda con "INFONAVIT", se procede a buscar en la columna B
If Not IMSS Is Nothing Then
    ' Obtiene la última fila ocupada en la columna B
    Set iRange = objWorksheetN.Range("B" & IMSS.Row & ":B" & objWorksheetN.Cells(objWorksheetN.Rows.Count, 2).End(-4162).Row)
    ' Itera desde la fila encontrada hasta la última fila ocupada en la columna 37 (columna AK)
    For i = lastAK4 to objWorksheetIMSS.Cells(objWorksheetIMSS.Rows.Count, 37).End(-4162).Row
        ' Si la celda en la columna 37 (columna AK) no está vacía
        If objWorksheetIMSS.Cells(i, 37).Value <> "" Then
            ' Busca el valor de la columna 37 (columna AK) en la primera columna del catálogo
            For j = 1 To ultimaFilaC
                If objWorksheetIMSS.Cells(i, 37).Value = objWorksheetCatalogo.Cells(j,1).value Then
                    ' Si se encuentra, se asigna el valor de la columna 2 del catálogo a CC
                    ' y se sale del bucle
                    CC = objWorksheetCatalogo.Cells(j,2).value
                    Exit For
                Else
                    ' Si no se encuentra, se asigna el valor de la columna 37 (columna AK) a CC
                    ' para buscarlo en la columna B de la hoja de Nomina
                    CC = objWorksheetIMSS.Cells(i, 37).Value
                End If
            Next
            ' Busca el valor de la columna 37 (columna AK) en el rango de la columna B
            Set iCC = iRange.Find(CC,,xlValues,xlPart)
            ' Si se encuentra el valor, se copia el valor de la columna 38 (columna AL) en la columna E de la hoja de Nomina
            If Not iCC Is Nothing Then
                objWorksheetN.Cells(iCC.Row, 5).Value = objWorksheetIMSS.Cells(i, 38).Value ' Columna E
				' Pega el UUID
				objWorksheetN.Cells(iCC.Row, 3).Value = CapturaUUID
            End If
        End If
    Next
    Set iCC = Nothing
End If

'Guarda y cierre el libro
objWorkbookIMSS.Save
objWorkbookIMSS.Close

objWorkbookN.Save
objWorkbookN.Close

objWorkbookCatalogo.Save
objWorkbookCatalogo.Close

'Quita la instancia del objeto Excel
objExcel.Quit

'Devuelve el error en caso de
If Err.Number <> 0 Then
	' Cerrar la aplicación de Excel
    objExcel.Quit
    Msg = "Error was generated by " & Err.Source & ". " & Err.Description
    WScript.StdOut.WriteLine Msg
Else
    WScript.StdOut.WriteLine "Script executed successfully."
End if