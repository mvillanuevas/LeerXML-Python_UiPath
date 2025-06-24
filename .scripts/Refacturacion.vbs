'On Error Resume Next

'Set objArgs = WScript.Arguments

'WorkbookPathRexmex = objArgs(0)
'WorkbookSheetRexmex = objArgs(1)
'WorkbookPathRef = objArgs(2)
'WorkbookSheetRef = objArgs(3)
'ActualMonth = objArgs(4)

WorkbookPathRexmex = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\REXMEX - Cuenta Operativa 2025_120525.xlsx"
WorkbookSheetRexmex = "Cuenta Operativa"
WorkbookPathRef = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\Layout refacturación may-25.xlsx"
ActualMonth = 3
WorkbookSheetLayout = "Layout"

'Genera un objeto de tipo Excel Application
Set objExcel = CreateObject("Excel.Application")

'Parámetro para indicar si se quiere visible la aplicación de Excel
objExcel.Application.Visible = True
'Evita movimiento de pantalla
objExcel.Application.ScreenUpdating = True
'Parámetro evitar mostrar pop ups de Excel
objExcel.Application.DisplayAlerts = False

'Abre libro Excel
Set objWorkbookPathRef = objExcel.Workbooks.Open(WorkbookPathRef, 0)
Set objWorkbookSheetRefL = objWorkbookPathRef.Worksheets(WorkbookSheetLayout)

Set objWorkbookPathRexmex = objExcel.Workbooks.Open(WorkbookPathRexmex, 0)
Set objWorkbookSheetRexmex = objWorkbookPathRexmex.Worksheets(WorkbookSheetRexmex)

' Arreglo de hojas de refacturación
Dim refacturacionSheets, bloque
refacturacionSheets = Array("BL29", "BL10", "BL11", "BL14")

' Iteraer sobre las hojas de refacturación
Dim i

For i = LBound(refacturacionSheets) To UBound(refacturacionSheets)
    Dim sheetName
    sheetName = refacturacionSheets(i)
    
    ' Verificar si la hoja existe en el libro de refacturación
    If SheetExists(objWorkbookPathRef, sheetName) Then
        Set objWorkbookSheetRef = objWorkbookPathRef.Worksheets(sheetName)

        ' Verificar si los filtros están activos en la fila 1, si no, activarlos
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


        ' Encontrar la última fila con datos en la columna a filtrar 
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

        ' Encontrar la última fila con datos en la hoja de Layout refacturación
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

' Encontrar la última fila con datos en la hoja de Layout refacturación
lastRowL = objWorkbookSheetRefL.Cells(objWorkbookSheetRefL.Rows.Count, 1).End(-4162).Row ' -4162 = xlUp

' Ocultar las filas que no cumplen con el criterio de la columna 1 (BL29)
Dim j
For j = 7 To lastRowL
    If objWorkbookSheetRefL.Cells(j, 1).Value <> "BLOQUE 29" Then
        objWorkbookSheetRefL.Rows(j).Hidden = True
    End If
Next
MsgBox "Se han copiado los datos de las hojas de refacturación a la hoja de Layout."
Set objWorkbookSheetRef = objWorkbookPathRef.Worksheets("BL29")
' Encontrar la última fila con datos en la columna a filtrar 
lastRow = objWorkbookSheetRef.Cells(objWorkbookSheetRefL.Rows.Count, 1).End(-4162).Row ' -4162 = xlUp

' Ordenar de manera ascendente la columna AG (columna 33) "UUID"
With objWorkbookSheetRef.Sort
    .SortFields.Clear
    .SortFields.Add objWorkbookSheetRef.Range("AG2:AG" & lastRow), 0, 1 ' 0 = xlSortOnValues, 1 = xlAscending
    .SetRange objWorkbookSheetRef.Range("A1:AV" & lastRow) ' Ajusta el rango según tus datos
    .Header = 1 ' 1 = xlYes (hay encabezado)
    .Apply
End With

' Revisar duplicados en columna C (WBS, col 3) y modificar columna AG (PEP, col 33)
Dim rowNum, wbsValue, agValue, sKey
Dim wbsDict, wbsDictH, pepCounterDict

Set wbsDict     = CreateObject("Scripting.Dictionary") ' Para detectar duplicados por UUID
Set wbsDictH    = CreateObject("Scripting.Dictionary") ' Para detectar duplicados por WBS+UUID
Set pepCounterDict = CreateObject("Scripting.Dictionary") ' Contador por UUID (o WBS)

For rowNum = 2 To lastRow ' Asumiendo encabezado en la fila 1
    wbsValue = Trim(objWorkbookSheetRef.Cells(rowNum, 3).Value)
    agValue  = Trim(objWorkbookSheetRef.Cells(rowNum, 33).Value)
    
    If wbsValue <> "" And agValue <> "" Then
        sKey = agValue & "|" & wbsValue ' Clave compuesta UUID + WBS

        ' Contador por UUID
        If Not wbsDict.Exists(agValue) Then
            wbsDict.Add agValue, 1
            pepCounterDict.Add agValue, 0
        Else
            pepCounterDict(agValue) = pepCounterDict(agValue) + 1
            objWorkbookSheetRef.Cells(rowNum, 33).Value = agValue & " pep" & pepCounterDict(agValue)
            objWorkbookSheetRef.Rows(rowNum).Hidden = True ' Descomenta si deseas ocultar
        End If

        ' Detección de duplicados exactos UUID+WBS
        If Not wbsDictH.Exists(sKey) Then
            wbsDictH.Add sKey, 1
        Else
            objWorkbookSheetRef.Rows(rowNum).Hidden = True ' Descomenta si deseas ocultar
        End If
    End If
Next

' Copiar columnas específicas de objWorkbookSheetRef a objWorkbookSheetRefL

Dim copyLastRow, pasteLastRow
copyLastRow = objWorkbookSheetRef.Cells(objWorkbookSheetRef.Rows.Count, 1).End(-4162).Row
pasteLastRow = objWorkbookSheetRefL.Cells(objWorkbookSheetRefL.Rows.Count, 1).End(-4162).Row + 1

' AP (col 42) -> D (col 4)
objWorkbookSheetRef.Range("AP2:AP" & copyLastRow).Copy
objWorkbookSheetRefL.Range("D" & pasteLastRow).PasteSpecial -4163

' AG (col 33) -> E (col 5)
objWorkbookSheetRef.Range("AG2:AG" & copyLastRow).Copy
objWorkbookSheetRefL.Range("E" & pasteLastRow).PasteSpecial -4163

' AI (col 35) -> I (col 9)
objWorkbookSheetRef.Range("AI2:AI" & copyLastRow).Copy
objWorkbookSheetRefL.Range("I" & pasteLastRow).PasteSpecial -4163

' AH (col 34) -> M (col 13)
objWorkbookSheetRef.Range("AH2:AH" & copyLastRow).Copy
objWorkbookSheetRefL.Range("M" & pasteLastRow).PasteSpecial -4163

' AE (col 31) -> R (col 18)
objWorkbookSheetRef.Range("AE2:AE" & copyLastRow).Copy
objWorkbookSheetRefL.Range("R" & pasteLastRow).PasteSpecial -4163

' L (col 12) -> V (col 22)
objWorkbookSheetRef.Range("L2:L" & copyLastRow).Copy
objWorkbookSheetRefL.Range("V" & pasteLastRow).PasteSpecial -4163

' F (col 6) -> X (col 24)
objWorkbookSheetRef.Range("F2:F" & copyLastRow).Copy
objWorkbookSheetRefL.Range("X" & pasteLastRow).PasteSpecial -4163

' G (col 7) -> Y (col 25)
objWorkbookSheetRef.Range("G2:G" & copyLastRow).Copy
objWorkbookSheetRefL.Range("Y" & pasteLastRow).PasteSpecial -4163

' H (col 8) -> Z (col 26)
objWorkbookSheetRef.Range("H2:H" & copyLastRow).Copy
objWorkbookSheetRefL.Range("Z" & pasteLastRow).PasteSpecial -4163

' I (col 9) -> AA (col 27)
objWorkbookSheetRef.Range("I2:I" & copyLastRow).Copy
objWorkbookSheetRefL.Range("AA" & pasteLastRow).PasteSpecial -4163

objExcel.CutCopyMode = False


' Función para validar si una hoja existe en un libro de Excel
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