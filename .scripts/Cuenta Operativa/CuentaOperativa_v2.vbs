WorkbookPathRexmex = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\REXMEX - Cuenta Operativa 2025_120525.xlsx"
WorkbookSheetRexmex = "Cuenta Operativa"
WorkbookPathN = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\Nomina Abr25_conNoAcreedor.xlsx"
WorkbookSheetNomina = "NOMINA"
anio = 2025
mes = 3

Set objExcel = CreateObject("Excel.Application")
objExcel.Application.Visible = True
objExcel.Application.ScreenUpdating = True
objExcel.Application.DisplayAlerts = False
objExcel.Application.EnableEvents = False

Set objWorkbookPathNomina = objExcel.Workbooks.Open(WorkbookPathN)
Set objWorkbookSheetNomina = objWorkbookPathNomina.Worksheets(WorkbookSheetNomina)
Set objWorkbookPathRexmex = objExcel.Workbooks.Open(WorkbookPathRexmex)
Set objWorkbookSheetRexmex = objWorkbookPathRexmex.Worksheets(WorkbookSheetRexmex)

If objWorkbookSheetNomina.AutoFilterMode Then objWorkbookSheetNomina.AutoFilterMode = False
If objWorkbookSheetRexmex.AutoFilterMode Then objWorkbookSheetRexmex.AutoFilterMode = False

ultimoDiaMes = DateSerial(anio, mes + 1, 0)

Set filesys = CreateObject("Scripting.FileSystemObject")
nombreArchivo = filesys.GetFileName(WorkbookPathN)
nombreArchivo = Replace(nombreArchivo, ".XLSX", "")
nombreArchivo = Replace(nombreArchivo, ".xlsx", "")

' Encontrar la última fila y columna a procesar
lastRow = objWorkbookSheetNomina.Cells(objWorkbookSheetNomina.Rows.Count, 2).End(-4162).Row
lastCol = objWorkbookSheetNomina.Cells(3, objWorkbookSheetNomina.Columns.Count).End(-4159).Column

' Leer todos los datos a memoria
Dim nominaData, rexData
nominaData = objWorkbookSheetNomina.Range(objWorkbookSheetNomina.Cells(1, 1), objWorkbookSheetNomina.Cells(lastRow, lastCol)).Value
rexLastRow = objWorkbookSheetRexmex.Cells(objWorkbookSheetRexmex.Rows.Count, 1).End(-4162).Row
' Asumimos que rexData tiene suficiente tamaño
ReDim rexData(10000, 70)  ' Ajusta 10000 según necesidad, o calcula dinámico

' Buscar primera fila de datos en columna A
For i = 10 To lastRow
    If Trim(nominaData(i, 1)) = "" Then
        firstRow = i + 1
        Exit For
    End If
Next

Set uniqueDict = CreateObject("Scripting.Dictionary")
Set pepDict = CreateObject("Scripting.Dictionary")

' Variables para guardar valores temporales
Dim totalFactura, valorAcreedor, numDocto, uuid, fecha, pep, valorPEP
rowR = 0
Dim rexOutRow
rexOutRow = rexLastRow

For i = firstRow To lastRow
    If nominaData(i, 2) = "" And nominaData(i, 3) = "" And nominaData(i, 5) = "" Then
        For j = 9 To lastCol
            If nominaData(i, j) <> "" And nominaData(i, j) <> 0 Then
                pep = nominaData(3, j)
                valorPEP = nominaData(i, j)
                rowR = rexOutRow + (j - 8)
                ' Llenar rexData en memoria
                rexData(rowR, 1) = "PEP"
                rexData(rowR, 2) = "FP"
                rexData(rowR, 3) = "MX29"
                rexData(rowR, 4) = anio
                rexData(rowR, 5) = mes
                rexData(rowR, 6) = ultimoDiaMes
                rexData(rowR, 7) = pep
                rexData(rowR, 16) = nombreArchivo
                rexData(rowR, 20) = valorAcreedor
                rexData(rowR, 36) = valorPEP
                rexData(rowR, 54) = totalFactura
                rexData(rowR, 56) = uuid
                rexData(rowR, 57) = fecha
                rexData(rowR, 60) = numDocto
            End If
        Next
        rexOutRow = rexOutRow + (lastCol - 8)
    ElseIf nominaData(i, 5) <> "" And nominaData(i, 7) <> "" Then
        totalFactura = nominaData(i, 5)
        valorAcreedor = nominaData(i, 1)
        numDocto = nominaData(i, 7)
    ElseIf nominaData(i, 3) <> "" And nominaData(i, 4) <> "" Then
        uuid = nominaData(i, 3)
        fecha = nominaData(i, 4)
    End If
Next

' Escribir los datos generados en bloque
If rexOutRow > rexLastRow Then
    objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(rexLastRow + 1, 1), objWorkbookSheetRexmex.Cells(rexOutRow, 70)).Value = _
        objExcel.Application.WorksheetFunction.Index(rexData, 0, 0)
End If

' Autofill de rangos (hazlos solo 1 vez y sobre el rango de datos agregados)
If rexOutRow > rexLastRow Then
    objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(2, 24), objWorkbookSheetRexmex.Cells(2, 31)).AutoFill _
        objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(2, 24), objWorkbookSheetRexmex.Cells(rexOutRow, 31))
    objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(2, 33), objWorkbookSheetRexmex.Cells(2, 35)).AutoFill _
        objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(2, 33), objWorkbookSheetRexmex.Cells(rexOutRow, 35))
    objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(2, 37), objWorkbookSheetRexmex.Cells(2, 51)).AutoFill _
        objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(2, 37), objWorkbookSheetRexmex.Cells(rexOutRow, 51))
    objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(2, 64), objWorkbookSheetRexmex.Cells(2, 66)).AutoFill _
        objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(2, 64), objWorkbookSheetRexmex.Cells(rexOutRow, 66))
    objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(2, 68), objWorkbookSheetRexmex.Cells(2, 70)).AutoFill _
        objWorkbookSheetRexmex.Range(objWorkbookSheetRexmex.Cells(2, 68), objWorkbookSheetRexmex.Cells(rexOutRow, 70))
End If

' Eliminar filas vacías solo una vez al final (opcional, según tu lógica)
' Recorre de abajo hacia arriba para eliminar filas vacías solo en la nueva sección

For j = rexOutRow To 2 Step -1
    If objWorkbookSheetRexmex.Cells(j, 1).Value = "" Then
        objWorkbookSheetRexmex.Rows(j).Delete
    End If
Next

objExcel.Application.ScreenUpdating = True
objExcel.Application.EnableEvents = True

' Limpieza
Set objWorkbookSheetNomina = Nothing
Set objWorkbookPathNomina = Nothing
Set objWorkbookSheetRexmex = Nothing
Set objWorkbookPathRexmex = Nothing
Set objExcel = Nothing