On Error Resume Next

Set objArgs = WScript.Arguments

WorkbookPathBanco = objArgs(0)
WorkbookSheetBanco = objArgs(1)
WorkbookPathNomina = objArgs(2)
WorkbookSheetNomina = objArgs(3)

'WorkbookPathBanco = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\Bancos 30.04.25- Aplicacion.xlsx"
'WorkbookSheetBanco = "Sheet1"
'WorkbookPathNomina = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\Nomina abr25.XLSX"
'WorkbookSheetNomina = "NOMINA"

'Genera un objeto de tipo Excel Application
Set objExcel = CreateObject("Excel.Application")

'Par�metro para indicar si se quiere visible la aplicaci�n de Excel
objExcel.Application.Visible = False
'Evita movimiento de pantalla
objExcel.Application.ScreenUpdating = False
'Par�metro evitar mostrar pop ups de Excel
objExcel.Application.DisplayAlerts = False

'Abre libro Excel
Set objWorkbookBanco = objExcel.Workbooks.Open(WorkbookPathBanco, 0) ' 0 = UpdateLinks None
Set objWorksheetBanco = objWorkbookBanco.Worksheets(WorkbookSheetBanco)

' Verificar si los filtros est�n activos en la fila 1, si no, activarlos
If objWorksheetBanco.AutoFilterMode Then
    objWorksheetBanco.Rows(1).AutoFilter
End If

If Not objWorksheetBanco.AutoFilterMode Then
    objWorksheetBanco.Rows(1).AutoFilter
End If

colToFilter = 1 ' N�mero de columna a filtrar (1 = columna A)
' Encontrar la �ltima fila con datos en la columna a filtrar
lastRow = objWorksheetBanco.Cells(objWorksheetBanco.Rows.Count, colToFilter).End(-4162).Row ' -4162 = xlUp

' Aplicar autofiltro con el valor "ISN"
objWorksheetBanco.Range(objWorksheetBanco.Cells(1, colToFilter), objWorksheetBanco.Cells(lastRow, colToFilter)).AutoFilter _
                        17, "*ISN*", 2 ' 2 = xlFilterValues (criterio de texto con comodines)

Dim dRange, dDict
' Recuperar las celdas visibles de la columna R (columna 18)
Set dRange = objWorksheetBanco.Range("R2:R" & lastRow).SpecialCells(12) ' 12 = xlCellTypeVisible
' Crear un diccionario para almacenar los valores �nicos de la columna R
Set dDict = CreateObject("Scripting.Dictionary")

' Recorrer las celdas visibles de la columna R (columna 18)
For Each cell In dRange
    ' Asegurarse de que la celda no est� vac�a antes de agregarla al diccionario
    If cell.Value <> "" Then
        ' Agregar el valor de la celda al diccionario
        dDict.Add cell.Value, True
    End If
Next

' Contar el numero de keys en el diccionario
Dim countKeys
' Esto solo si existe mas de un resultado del filtro anterior
' Se puede incluir cada resultado en la inserci�n en la hoja de Nomina
countKeys = dDict.Count

'Abre libro Excel
Set objWorkbookN = objExcel.Workbooks.Open(WorkbookPathNomina)
Set objWorksheetN = objWorkbookN.Worksheets(WorkbookSheetNomina)

Const xlPart = 2
Const xlValues = -4163

' Establece rango de busqueda en la columna C
Set nRange = objWorksheetN.Range("C:C")
' Busca la celda que contiene "ISN" en la columna C
Dim ISN : Set ISN = nRange.Find("ISN",,xlValues,xlPart)

If Not ISN Is Nothing Then
    ' Iteramos 3 veces para insertar el valor de la llave en las siguientes 3 filas
    For i = 1 To 3
        ' Insertar el valor de la llave en la columna C, fila ISN.Row + i
        objWorksheetN.Cells((ISN.Row + 1) + i, 3).Value = dDict.Keys()(0)
    Next
End If
WScript.Echo dDict.Keys()(0)
' Guardar y cerrar el libro de Excel
objWorkbookBanco.Save
objWorkbookBanco.Close

' Guardar y cerrar el libro de Excel
objWorkbookN.Save
objWorkbookN.Close

' Cerrar la aplicaci�n de Excel
objExcel.Quit

'Devuelve el error en caso de
If Err.Number <> 0 Then
    Msg = "Error was generated by " & Err.Source & ". " & Err.Description
    WScript.Echo Msg
End if