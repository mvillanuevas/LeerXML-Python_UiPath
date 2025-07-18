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
WorkbookPathN = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\Nomina abr25_NoDocumentos.XLSX"
WorkbookSheetNomina = "NOMINA"
anio = "2025"
mes = "3"

'Genera un objeto de tipo Excel Application
Set objExcel = CreateObject("Excel.Application")

'Par�metro para indicar si se quiere visible la aplicaci�n de Excel
objExcel.Application.Visible = True
'Evita movimiento de pantalla
objExcel.Application.ScreenUpdating = True
'Par�metro evitar mostrar pop ups de Excel
objExcel.Application.DisplayAlerts = False

'Abre libro Excel
Set WorkbookPathNomina = objExcel.Workbooks.Open(WorkbookPathN)
Set WorkbookSheetNomina = WorkbookPathNomina.Worksheets(WorkbookSheetNomina)

Set WorkbookPathRexmex = objExcel.Workbooks.Open(WorkbookPathRexmex)
Set WorkbookSheetRexmex = WorkbookPathRexmex.Worksheets(WorkbookSheetRexmex)

' Verificar si los filtros est�n activos en la fila 1, si no, activarlos
If WorkbookSheetNomina.AutoFilterMode Then
    WorkbookSheetNomina.Rows(10).AutoFilter
End If

If Not WorkbookSheetNomina.AutoFilterMode Then
    WorkbookSheetNomina.Rows(10).AutoFilter
End If

dim filesys
Set filesys = CreateObject("Scripting.FileSystemObject")
nombreArchivo = filesys.GetFileName(WorkbookPathN)

nombreArchivo = Replace(nombreArchivo, ".XLSX", "")
nombreArchivo = Replace(nombreArchivo, ".xlsx", "")

' Encontrar la �ltima fila con datos en la columna a filtrar
lastRow = WorkbookSheetNomina.Cells(WorkbookSheetNomina.Rows.Count, 1).End(-4162).Row ' -4162 = xlUp
lastCol = WorkbookSheetNomina.Cells(3, WorkbookSheetNomina.Columns.Count).End(-4159).Column ' -4159 = xlToLeft

' Iterar la columna A hasta encontrar una celda vacia
For i = 10 to lastRow
    If WorkbookSheetNomina.Cells(i, 1).Value = "" Then
        lastRow = i - 1
        Exit For
    End If
Next

Dim ultimoDiaMes
ultimoDiaMes = DateSerial(anio, mes + 1, 0)

' Columna inical
colToFilter = 8 ' N�mero de columna a filtrar (1 = columna A)

For i = 9 To lastCol
    rowCount = 0
    ' Obtiene el valor de pep
    pep = WorkbookSheetNomina.Cells(3, i).Value
    ' Aplicar autofiltro con el valor "ISN"' Aplicar autofiltro con el valor diferente de "-"
    WorkbookSheetNomina.Range(WorkbookSheetNomina.Cells(10, i), WorkbookSheetNomina.Cells(lastRow, i)).AutoFilter _
                              i, "<>0"
    On Error Resume Next
    ' Recuperar las celdas visibles de la columna filtrada
    Set dRange = WorkbookSheetNomina.Range(WorkbookSheetNomina.Cells(11, i), WorkbookSheetNomina.Cells(lastRow, i)).SpecialCells(12) ' 12 = xlCellTypeVisible
    Set uRange = WorkbookSheetNomina.Range(WorkbookSheetNomina.Cells(11, 3), WorkbookSheetNomina.Cells(lastRow, 3)).SpecialCells(12) ' 12 = xlCellTypeVisible
    Set fRange = WorkbookSheetNomina.Range(WorkbookSheetNomina.Cells(11, 4), WorkbookSheetNomina.Cells(lastRow, 4)).SpecialCells(12) ' 12 = xlCellTypeVisible
    Set mRange = WorkbookSheetNomina.Range(WorkbookSheetNomina.Cells(11, 5), WorkbookSheetNomina.Cells(lastRow, 5)).SpecialCells(12) ' 12 = xlCellTypeVisible
    Set iRange = WorkbookSheetNomina.Range(WorkbookSheetNomina.Cells(11, 6), WorkbookSheetNomina.Cells(lastRow, 6)).SpecialCells(12) ' 12 = xlCellTypeVisible
    ' Contar las celdas visibles
    rowCount = dRange.Count

    ' Verificar si hay celdas visibles y si no hay error
    If rowCount > 0 And Err.Number = 0 Then
        ' Obtener la �ltima fila con datos en la hoja de REXMEX
        lastRowR = WorkbookSheetRexmex.Cells(WorkbookSheetRexmex.Rows.Count, 1).End(-4162).Row ' -4162 = xlUp
        ' Llenar las filas de la hoja de REXMEX con los datos de la hoja de NOMINA
        For j = 1 To rowCount
            WorkbookSheetRexmex.Rows(WorkbookSheetRexmex.Cells(WorkbookSheetRexmex.Rows.Count, 1).End(-4162).Row).Copy
            WorkbookSheetRexmex.Rows(lastRowR + j).Insert -4121 ' -4121 = xlDown
            WorkbookSheetRexmex.Cells(lastRowR + j, 4).Value = anio
            WorkbookSheetRexmex.Cells(lastRowR + j, 5).Value = mes
            WorkbookSheetRexmex.Cells(lastRowR + j, 5).Value = mes
            WorkbookSheetRexmex.Cells(lastRowR + j, 6).Value = ultimoDiaMes
            WorkbookSheetRexmex.Cells(lastRowR + j, 7).Value = pep
            WorkbookSheetRexmex.Cells(lastRowR + j, 16).Value = nombreArchivo
            'WorkbookSheetRexmex.Cells(lastRowR + j, 34).Value = "Repsol Exploracion Mexico, S.A de C.V."
            WorkbookSheetRexmex.Cells(lastRowR + j, 61).Value = ""
			
			WorkbookSheetRexmex.Cells(lastRowR + j, 33).Value = ""
			WorkbookSheetRexmex.Cells(lastRowR + j, 37).Value = ""
			WorkbookSheetRexmex.Cells(lastRowR + j, 55).Value = ""
        Next
        ' Copiar los valores de las celdas visibles a la hoja de REXMEX
        uRange.Copy
        WorkbookSheetRexmex.Range("BD" & lastRowR + 1).PasteSpecial -4163 ' -4163 = xlPasteAll
        fRange.Copy
        WorkbookSheetRexmex.Range("BE" & lastRowR + 1).PasteSpecial -4163 ' -4163 = xlPasteAll
        mRange.Copy
        WorkbookSheetRexmex.Range("BB" & lastRowR + 1).PasteSpecial -4163 ' -4163 = xlPasteAll
        iRange.Copy
        WorkbookSheetRexmex.Range("AZ" & lastRowR + 1).PasteSpecial -4163 ' -4163 = xlPasteAll
        dRange.Copy
        WorkbookSheetRexmex.Range("AJ" & lastRowR + 1).PasteSpecial -4163 ' -4163 = xlPasteAll
        ' Quitar el modo de corte/copia
        objExcel.CutCopyMode = False
        'MsgBox "Se han copiado " & rowCount & " filas de la columna " & i & " con el valor " & pep & " a la hoja de REXMEX."
    End If
    Err.Clear
    WorkbookSheetNomina.ShowAllData
Next

' Guardar y cerrar los libros de trabajo
'WorkbookPathRexmex.Save
'WorkbookPathRexmex.Close

WorkbookPathNomina.Save
WorkbookPathNomina.Close
' Cerrar la aplicaci�n de Excel
'objExcel.Quit

'Devuelve el error en caso de
If Err.Number <> 0 Then
    Msg = "Error was generated by " & Err.Source & ". " & Err.Description
    WScript.StdOut.WriteLine Msg
End if
