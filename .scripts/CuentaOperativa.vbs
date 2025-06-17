'On Error Resume Next

'Set objArgs = WScript.Arguments

'WorkbookPathRexmex = objArgs(0)
'WorkbookSheetRexmex = objArgs(1)
'WorkbookPathNomina = objArgs(2)
'WorkbookSheetNomina = objArgs(3)

WorkbookPathRexmex = "C:\Users\se109874\OneDrive - Repsol\Archivos de flores vega, yazmin (ext) - Reporte Regulatorio RPA\4 - Abril\Files\REXMEX - Cuenta Operativa 2025_120525.xlsx"
WorkbookSheetRexmex = "Cuenta Operativa"
WorkbookPathNomina = "C:\Users\se109874\OneDrive - Repsol\Archivos de flores vega, yazmin (ext) - Reporte Regulatorio RPA\4 - Abril\Files\Nomina abr25.XLSX"
WorkbookSheetNomina = "NOMINA"

'Genera un objeto de tipo Excel Application
Set objExcel = CreateObject("Excel.Application")

'Parámetro para indicar si se quiere visible la aplicación de Excel
objExcel.Application.Visible = False
'Evita movimiento de pantalla
objExcel.Application.ScreenUpdating = False
'Parámetro evitar mostrar pop ups de Excel
objExcel.Application.DisplayAlerts = False

'Abre libro Excel
Set objWorkbookRexmex = objExcel.Workbooks.Open(WorkbookPathNomina)
Set objWorksheetRexmex = objWorkbookRexmex.Worksheets(WorkbookSheetNomina)

' Verificar si los filtros están activos en la fila 1, si no, activarlos
If objWorksheetBanco.AutoFilterMode Then
    objWorksheetBanco.Rows(1).AutoFilter
End If

If Not objWorksheetBanco.AutoFilterMode Then
    objWorksheetBanco.Rows(1).AutoFilter
End If

colToFilter = 1 ' Número de columna a filtrar (1 = columna A)
' Encontrar la última fila con datos en la columna a filtrar
lastRow = objWorksheetBanco.Cells(objWorksheetBanco.Rows.Count, colToFilter).End(-4162).Row ' -4162 = xlUp

' Aplicar autofiltro con el valor "ISN"
objWorksheetRexmex.Range(objWorksheetRexmex.Cells(1, colToFilter), objWorksheetBanco.Cells(lastRow, colToFilter)).AutoFilter _
                        17, "-", 7  ' 7 = xlFilterValues, operador "diferente de"

