'On Error Resume Next

'Set objArgs = WScript.Arguments

'WorkbookPathRexmex = objArgs(0)
'WorkbookSheetRexmex = objArgs(1)
'WorkbookPathRef = objArgs(2)
'WorkbookSheetRef = objArgs(3)
'ActualMonth = objArgs(4)

WorkbookPathRexmex = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\REXMEX - Cuenta Operativa 2025_120525.xlsx"
WorkbookSheetRexmex = "Cuenta Operativa"
'WorkbookPathRef = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\Nomina abr25.XLSX"
'WorkbookSheetRef = "NOMINA"
ActualMonth = 3

'Genera un objeto de tipo Excel Application
Set objExcel = CreateObject("Excel.Application")

'Parámetro para indicar si se quiere visible la aplicación de Excel
objExcel.Application.Visible = True
'Evita movimiento de pantalla
objExcel.Application.ScreenUpdating = True
'Parámetro evitar mostrar pop ups de Excel
objExcel.Application.DisplayAlerts = False

'Abre libro Excel
'Set objWorkbookPathRef = objExcel.Workbooks.Open(WorkbookPathRef)
'Set objWorkbookSheetRef = objWorkbookPathNomina.Worksheets(WorkbookSheetRef)

Set objWorkbookPathRexmex = objExcel.Workbooks.Open(WorkbookPathRexmex)
Set objWorkbookSheetRexmex = objWorkbookPathRexmex.Worksheets(WorkbookSheetRexmex)

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

MsgBox "Filtró aplicado desde " & primerDiaMes & " hasta " & ultimoDiaMes
