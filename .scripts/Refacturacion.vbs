'On Error Resume Next

'Set objArgs = WScript.Arguments

'WorkbookPathRexmex = objArgs(0)
'WorkbookSheetRexmex = objArgs(1)
'WorkbookPathRef = objArgs(2)
'WorkbookSheetRef = objArgs(3)
'ActualMonth = objArgs(4)

WorkbookPathRexmex = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\REXMEX - Cuenta Operativa 2025_120525.xlsx"
WorkbookPathRef = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\Layout refacturaciùn may-25.xlsx"
ActualMonth = 3

WorkbookSheetRexmex = "Cuenta Operativa"
WorkbookSheetLayout = "Layout"

'Genera un objeto de tipo Excel Application
Set objExcel = CreateObject("Excel.Application")

'Parùmetro para indicar si se quiere visible la aplicaciùn de Excel
objExcel.Application.Visible = True
'Evita movimiento de pantalla
objExcel.Application.ScreenUpdating = True
'Parùmetro evitar mostrar pop ups de Excel
objExcel.Application.DisplayAlerts = False

'Abre libro Excel
Set objWorkbookPathRef = objExcel.Workbooks.Open(WorkbookPathRef, 0)
Set objWorkbookSheetRefL = objWorkbookPathRef.Worksheets(WorkbookSheetLayout)

Set objWorkbookPathRexmex = objExcel.Workbooks.Open(WorkbookPathRexmex, 0)
Set objWorkbookSheetRexmex = objWorkbookPathRexmex.Worksheets(WorkbookSheetRexmex)

' Arreglo de hojas de refacturaciùn
Dim refacturacionSheets, bloque
refacturacionSheets = Array("BL29", "BL10", "BL11", "BL14")

' Iteraer sobre las hojas de refacturaciùn
Dim i

For i = LBound(refacturacionSheets) To UBound(refacturacionSheets)
    Dim sheetName
    sheetName = refacturacionSheets(i)
    
    ' Verificar si la hoja existe en el libro de refacturaciùn
    If SheetExists(objWorkbookPathRef, sheetName) Then
        Set objWorkbookSheetRef = objWorkbookPathRef.Worksheets(sheetName)

        ' Verificar si los filtros estùn activos en la fila 1, si no, activarlos
        If objWorkbookSheetRexmex.AutoFilterMode Then
            objWorkbookSheetRexmex.AutoFilterMode = False
        End If

        If Not objWorkbookSheetRexmex.AutoFilterMode Then
            objWorkbookSheetRexmex.Rows(1).AutoFilter
        End If

        Dim ultimoDiaMes
        ultimoDiaMes = DateSerial(Year(Date), ActualMonth + 1, 0)
        ultimoDiaMes = Right("0" & Month(ultimoDiaMes),2) & "-" & Right("0" & Day(ultimoDiaMes),2) & "-" & Year(ultimoDiaMes)

        primerDiaMes = DateSerial(Year(Date), ActualMonth, 1)
        primerDiaMes = Right("0" & Month(primerDiaMes),2) & "-" & Right("0" & Day(primerDiaMes),2) & "-" & Year(primerDiaMes)


        ' Encontrar la ùltima fila con datos en la columna a filtrar 
        lastRow = objWorkbookSheetRexmex.Cells(objWorkbookSheetRexmex.Rows.Count, 1).End(-4162).Row ' -4162 = xlUp

        objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(1, 27), objWorkbookSheetRexmex.Cells(lastRow, 27)).AutoFilter _
                                    27, ">=" & CDbl(CDate(primerDiaMes)), 1, "<=" & CDbl(CDate(ultimoDiaMes)), 1

        objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(1, 34), objWorkbookSheetRexmex.Cells(lastRow, 34)).AutoFilter _
                                    34, "<>OVERHEAD"

        objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(1, 55), objWorkbookSheetRexmex.Cells(lastRow, 55)).AutoFilter _
                                    55, "<>NC"

        objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(1, 24), objWorkbookSheetRexmex.Cells(lastRow, 24)).AutoFilter _
                                    24, "=" & sheetName
        
        Set dRange = objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(2, 24), objWorkbookSheetRexmex.Cells(lastRow, 71)).SpecialCells(12)

        ' Encontrar la ùltima fila con datos en la hoja de Layout refacturaciùn
        lastRowR = objWorkbookSheetRef.Cells(objWorkbookSheetRef.Rows.Count, 1).End(-4162).Row ' -4162 = xlUp

        ' Copiar los valores de las celdas visibles a la hoja de REXMEX
        dRange.Copy

        ' Mostrar todas las filas de la hoja (quitar ocultamiento)
        objWorkbookSheetRef.Rows.Hidden = False
        ' Mostrar todas las columnas de la hoja (quitar ocultamiento)
        objWorkbookSheetRef.Columns.Hidden = False

        objWorkbookSheetRef.Range("A" & lastRowR + 1).PasteSpecial -4163 ' -4163 = xlPasteAll
        ' Quitar el modo de corte/copia
        objExcel.CutCopyMode = False
    End If
Next

' Encontrar la ùltima fila con datos en la hoja de Layout refacturaciùn
lastRowL = objWorkbookSheetRefL.Cells(objWorkbookSheetRefL.Rows.Count, 1).End(-4162).Row ' -4162 = xlUp

' Ocultar las filas que no cumplen con el criterio de la columna 1 (BL29)
Dim j
For j = 7 To lastRowL
    If objWorkbookSheetRefL.Cells(j, 1).Value <> "BLOQUE 29" Then
        objWorkbookSheetRefL.Rows(j).Hidden = True
    End If
Next

Set objWorkbookSheetRef = objWorkbookPathRef.Worksheets("BL29")
' Encontrar la ùltima fila con datos en la columna a filtrar 
lastRow = objWorkbookSheetRef.Cells(objWorkbookSheetRef.Rows.Count, 1).End(-4162).Row ' -4162 = xlUp
MsgBox "ùltima fila con datos en BL29: " & lastRow
'Aplica Text to columns en formate General
objWorkbookSheetRef.Range("AG:AG").TextToColumns
' Ordenar de manera ascendente la columna AG (columna 33) "UUID"
With objWorkbookSheetRef.Sort
    .SortFields.Clear
    .SortFields.Add objWorkbookSheetRef.Range("AG2:AG" & lastRow), 0, 1 ' 0 = xlSortOnValues, 1 = xlAscending
    .SetRange objWorkbookSheetRef.Range("A1:AV" & lastRow) ' Ajusta el rango segùn tus datos
    .Header = 1 ' 1 = xlYes (hay encabezado)
    .Apply
End With

'____________________________________________________________________________________________________________________________________________

' --- Optimizaciùn de rendimiento: lectura y escritura por lotes ---
Dim wbsDict, wbsDictH, pepCounterDict
Set wbsDict     = CreateObject("Scripting.Dictionary")
Set wbsDictH    = CreateObject("Scripting.Dictionary")
Set pepCounterDict = CreateObject("Scripting.Dictionary")

Dim data, rowNum, totalRows
Dim wbsValue, agValue, sKey
Dim results()

' Leer todas las filas a un array (mùs rùpido que trabajar directo con Cells)
data = objWorkbookSheetRef.Range("A2:AV" & lastRow).Value ' A = col 1, AG = col 33
totalRows = UBound(data, 1)
ReDim results(totalRows) ' Guardar filas que se deben ocultar

For rowNum = 1 To totalRows ' Ya que empezamos en A2, este es ùndice 1
    wbsValue = Trim(data(rowNum, 3))     ' Columna C
    agValue  = Trim(data(rowNum, 33))    ' Columna AG
    
    If wbsValue <> "" And agValue <> "" Then
        sKey = agValue & "|" & wbsValue

        ' Contador por UUID
        If Not wbsDict.Exists(agValue) Then
            wbsDict.Add agValue, 1
            pepCounterDict.Add agValue, 0
        Else
            pepCounterDict(agValue) = pepCounterDict(agValue) + 1
            data(rowNum, 33) = agValue & " pep" & pepCounterDict(agValue)
            results(rowNum) = True ' Marcar para ocultar
        End If

        ' Duplicado exacto UUID + WBS
        If wbsDictH.Exists(sKey) Then
            results(rowNum) = True ' Marcar para ocultar
        Else
            wbsDictH.Add sKey, 1
        End If
    End If
Next

' Escribir los datos modificados de vuelta
objWorkbookSheetRef.Range("A2:AG" & lastRow).Value = data

' Ocultar filas en un solo paso
For rowNum = 1 To totalRows
    If results(rowNum) = True Then
        objWorkbookSheetRef.Rows(rowNum + 1).Hidden = True ' +1 por offset a partir de fila 2
    End If
Next

'____________________________________________________________________________________________________________________________________________

' Arreglo de hojas de refacturaciùn
Dim proveedores
proveedores = Array("PC CARIGALI", "PTTEP", "REPSOL", "SIERRA NEVADA")

saveLastRow = objWorkbookSheetRefL.Cells(objWorkbookSheetRefL.Rows.Count, 4).End(-4162).Row + 1

For i = LBound(proveedores) To UBound(proveedores)
    ' Copiar columnas especùficas de objWorkbookSheetRef a objWorkbookSheetRefL

    Dim copyLastRow, pasteLastRow
    copyLastRow = objWorkbookSheetRef.Cells(objWorkbookSheetRef.Rows.Count, 1).End(-4162).Row
    pasteLastRow = objWorkbookSheetRefL.Cells(objWorkbookSheetRefL.Rows.Count, 4).End(-4162).Row + 2

    ' AP (col 42) -> D (col 4)
    objWorkbookSheetRef.Range("AP2:AP" & copyLastRow).SpecialCells(12).Copy
    objWorkbookSheetRefL.Range("D" & pasteLastRow).PasteSpecial -4163

    ' AG (col 33) -> E (col 5)
    objWorkbookSheetRef.Range("AG2:AG" & copyLastRow).SpecialCells(12).Copy
    objWorkbookSheetRefL.Range("E" & pasteLastRow).PasteSpecial -4163
    
    'Iterar la columna E y validar si la longitud del valor del la celda es menor a 16 y si es asi, cortar los valores hacia la columna F
    Dim cell, longcell
    For Each cell In objWorkbookSheetRefL.Range("E" & pasteLastRow & ":E" & pasteLastRow + copyLastRow - 2)
        ' Si el valor de la celda contiene el valor "pep" restar 3 a la longitud del valor de la celda
        If InStr(1, cell.Value, "pep", vbTextCompare) > 0 Then
            longcell = Left(cell.Value) - 3
        Else
            longcell = Len(cell.Value)
        End If
        If longcell < 16 Then
            cell.Offset(0, 1).Value = cell.Value ' Mover el valor a la columna F
            cell.Value = "" ' Limpiar la celda original
        End If
    Next

    ' B (col 2) -> L (col 12)
    RowCount = objWorkbookSheetRefL.Cells(objWorkbookSheetRefL.Rows.Count, 4).End(-4162).Row
    
    objWorkbookSheetRefL.Range("E" & pasteLastRow & ":E" & RowCount).SpecialCells(12).Copy
    objWorkbookSheetRefL.Range("L" & pasteLastRow).PasteSpecial -4163

    ' AI (col 35) -> I (col 9)
    objWorkbookSheetRef.Range("AI2:AI" & copyLastRow).SpecialCells(12).Copy
    objWorkbookSheetRefL.Range("I" & pasteLastRow).PasteSpecial -4163

    ' AH (col 34) -> M (col 13)
    objWorkbookSheetRef.Range("AH2:AH" & copyLastRow).SpecialCells(12).Copy
    objWorkbookSheetRefL.Range("M" & pasteLastRow).PasteSpecial -4163

    ' N (col 14) -> O (col 15)
    objWorkbookSheetRef.Range("N2:N" & copyLastRow).SpecialCells(12).Copy
    objWorkbookSheetRefL.Range("O" & pasteLastRow).PasteSpecial -4163

    ' AE (col 31) -> R (col 18)
    objWorkbookSheetRef.Range("AE2:AE" & copyLastRow).SpecialCells(12).Copy
    objWorkbookSheetRefL.Range("R" & pasteLastRow).PasteSpecial -4163

    ' L (col 12) -> V (col 22)
    objWorkbookSheetRef.Range("L2:L" & copyLastRow).SpecialCells(12).Copy
    objWorkbookSheetRefL.Range("V" & pasteLastRow).PasteSpecial -4163

    ' F (col 6) -> X (col 24)
    objWorkbookSheetRef.Range("F2:F" & copyLastRow).SpecialCells(12).Copy
    objWorkbookSheetRefL.Range("X" & pasteLastRow).PasteSpecial -4163

    ' G (col 7) -> Y (col 25)
    objWorkbookSheetRef.Range("G2:G" & copyLastRow).SpecialCells(12).Copy
    objWorkbookSheetRefL.Range("Y" & pasteLastRow).PasteSpecial -4163

    ' H (col 8) -> Z (col 26)
    objWorkbookSheetRef.Range("H2:H" & copyLastRow).SpecialCells(12).Copy
    objWorkbookSheetRefL.Range("Z" & pasteLastRow).PasteSpecial -4163

    ' I (col 9) -> AA (col 27)
    objWorkbookSheetRef.Range("I2:I" & copyLastRow).SpecialCells(12).Copy
    objWorkbookSheetRefL.Range("AA" & pasteLastRow).PasteSpecial -4163

    ' C (col 3) -> AG (col 33)
    objWorkbookSheetRef.Range("C2:C" & copyLastRow).SpecialCells(12).Copy
    objWorkbookSheetRefL.Range("AG" & pasteLastRow).PasteSpecial -4163

    ' D (col 4) -> AH (col 34)
    objWorkbookSheetRef.Range("D2:D" & copyLastRow).SpecialCells(12).Copy
    objWorkbookSheetRefL.Range("AH" & pasteLastRow).PasteSpecial -4163

    ' E (col 5) -> AI (col 35)
    objWorkbookSheetRef.Range("E2:E" & copyLastRow).SpecialCells(12).Copy
    objWorkbookSheetRefL.Range("AI" & pasteLastRow).PasteSpecial -4163

    ' K (col 11) -> AJ (col 36)
    objWorkbookSheetRef.Range("K2:K" & copyLastRow).SpecialCells(12).Copy
    objWorkbookSheetRefL.Range("AJ" & pasteLastRow).PasteSpecial -4163

    ' AO (col 41) -> AK (col 37)
    objWorkbookSheetRef.Range("AO2:AO" & copyLastRow).SpecialCells(12).Copy
    objWorkbookSheetRefL.Range("AK" & pasteLastRow).PasteSpecial -4163

    objExcel.CutCopyMode = False

    ' Realizar autofill de fùrmulas en las columnas S, T, U, AD, AE de objWorkbookSheetRefL
    
    Dim fillLastRow
    fillLastRow = objWorkbookSheetRefL.Cells(objWorkbookSheetRefL.Rows.Count, 4).End(-4162).Row

    ' S (col 19)
    objWorkbookSheetRefL.Range("S7").AutoFill objWorkbookSheetRefL.Range("S7:S" & fillLastRow)
    ' T (col 20)
    objWorkbookSheetRefL.Range("T7").AutoFill objWorkbookSheetRefL.Range("T7:T" & fillLastRow)
    ' U (col 21)
    objWorkbookSheetRefL.Range("U7").AutoFill objWorkbookSheetRefL.Range("U7:U" & fillLastRow)
    ' Q (col 17)
    objWorkbookSheetRefL.Range("Q7").AutoFill objWorkbookSheetRefL.Range("Q7:Q" & fillLastRow)
    ' W (col 23s)
    objWorkbookSheetRefL.Range("W7").AutoFill objWorkbookSheetRefL.Range("W7:W" & fillLastRow)
    ' AD (col 30)
    objWorkbookSheetRefL.Range("AD7").AutoFill objWorkbookSheetRefL.Range("AD7:AD" & fillLastRow)
    ' AE (col 31)
    objWorkbookSheetRefL.Range("AE7").AutoFill objWorkbookSheetRefL.Range("AE7:AE" & fillLastRow)

    ' Limpiar una fila vacùa antes de pegar los datos
    objWorkbookSheetRefL.Rows(pasteLastRow - 1).ClearContents

    ' Rellenar con autofill el valor REP en la columna B de objWorkbookSheetRefL
    Dim bStart, bEnd
    bStart = pasteLastRow
    bEnd = fillLastRow

    ' Rellenar con autofill el valor "BLOQUE 29" en la columna A de objWorkbookSheetRefL
    objWorkbookSheetRefL.Range("A" & bStart).Value = "BLOQUE 29"
    objWorkbookSheetRefL.Range("A" & bStart & ":A" & bEnd).Value = "BLOQUE 29"
    ' Rellenar con autofill el valor del proveedor actual en la columna B de objWorkbookSheetRefL
    objWorkbookSheetRefL.Range("B" & bStart).Value = proveedores(i)
    objWorkbookSheetRefL.Range("B" & bStart & ":B" & bEnd).Value = proveedores(i)

    ' Rellenar con autofill el valor "Bloque 29, AP-CS-G10, Cuenca Salina / Administraciùn General" en la columna AC de objWorkbookSheetRefL
    objWorkbookSheetRefL.Range("AC" & bStart).Value = "Bloque 29, AP-CS-G10, Cuenca Salina / Administraciùn General"
    objWorkbookSheetRefL.Range("AC" & bStart & ":AC" & bEnd).Value = "Bloque 29, AP-CS-G10, Cuenca Salina / Administraciùn General"
    
Next

'____________________________________________________________________________________________________________________________________________

' Encontrar la ùltima fila con datos en la hoja de Layout refacturaciùn
LastRow = objWorkbookSheetRefL.Cells(objWorkbookSheetRefL.Rows.Count, 4).End(-4162).Row

' Aplicar negritas a un rango especùfico
With objWorkbookSheetRefL.Range("A" & saveLastRow & ":B" & LastRow)
    .Font.Bold = True
End With

' Aplicar All borders a un rago especùfico
With objWorkbookSheetRefL.Range("C" & saveLastRow & ":AF" & LastRow)
    .Borders.LineStyle = 1 ' xlContinuous
    .Borders.Weight = 2 ' xlMedium
End With

' Crear un arreglo de columnas para aplicar el formato Right border
Dim rightBorderCols
rightBorderCols = Array("C", "D", "E", "K", "M", "N", "V", "W", "Z", "AC", "AF")
' Aplicar Right border a las columnas especificadas
Dim col
For Each col In rightBorderCols
    With objWorkbookSheetRefL.Range(col & saveLastRow & ":" & col & LastRow)
        .Borders(10).LineStyle = 1 ' xlContinuous
        .Borders(10).Weight = -4138 ' xlMedium
    End With
Next
' Aplicar color de fondo a un rango especùfico
With objWorkbookSheetRefL.Range("C" & saveLastRow & ":AF" & LastRow)
    .Interior.Color = RGB(217, 225, 242) ' Color azul claro
End With

' Crear un nombre de hoja basado en la fecha y hora actual
sheetName = "Layout " & Day(Now) & Month(Now) & Year(Now) & "_" & Second(Now)
' Crear una copia de la hoja actual sobre el mismo libro
If Not SheetExists(objWorkbookPathRef, sheetName) Then
    objWorkbookSheetRefL.Copy objWorkbookPathRef.Sheets(objWorkbookPathRef.Sheets.Count)
    objWorkbookPathRef.Sheets(objWorkbookPathRef.Sheets.Count).Name = sheetName
Else
    objWorkbookPathRef.Sheets(sheetName).Delete
    objWorkbookSheetRefL.Copy objWorkbookPathRef.Sheets(objWorkbookPathRef.Sheets.Count)
    objWorkbookPathRef.Sheets(objWorkbookPathRef.Sheets.Count).Name = sheetName
End If

Set objWorkbookSheetRefLN = objWorkbookPathRef.Worksheets(sheetName)

' Copiar y pegar como valores todas las celdas de una hoja
With objWorkbookSheetRefLN.UsedRange
    .Copy
    .PasteSpecial -4163 ' xlPasteValues
End With
objExcel.CutCopyMode = False

MsgBox "Proceso de refacturaciùn completado."

'____________________________________________________________________________________________________________________________________________
' Funciùn para validar si una hoja existe en un libro de Excel
Function SheetExists(wb, sheetName)
    Dim s
    SheetExists = False
    For Each s In wb.Sheets
        If StrComp(s.Name, sheetName, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next
End Function