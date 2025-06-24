On Error Resume Next

Set objArgs = WScript.Arguments

WorkbookPathNomina = objArgs(0)
WorkbookSheetNomina = objArgs(1)
WorkbookPathReclas = objArgs(2)
WorkbookSheetReclas = objArgs(3)
WorkbookPathCatalogo = objArgs(4)
WorkbookSheetCatalogo = objArgs(5)

'WorkbookPathNomina = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\Nomina abr25.XLSX"
'WorkbookSheetNomina = "NOMINA"
'WorkbookPathReclas = "C:\Users\HE678HU\OneDrive - EY\.Repsol\Reporte Regulatorio\4 - Abril\Files\02 Reclas proveedores nomina (1).xlsx"
'WorkbookSheetReclas = "determinaciÛn"
'WorkbookPathCatalogo = "C:\Users\HE678HU\OneDrive - EY\Documents\UiPath\Leer_Facturas_Nomina\.templates\Catalogo Bloques.xlsx"
'WorkbookSheetCatalogo = "Centro"

'Genera un objeto de tipo Excel Application
Set objExcel = CreateObject("Excel.Application")
'Genera objeto diccionario
Set objDictionary = CreateObject("Scripting.Dictionary")
Set objDictionaryPorcentual = CreateObject("Scripting.Dictionary")

'Parùmetro para indicar si se quiere visible la aplicaciùn de Excel
objExcel.Application.Visible = True
'Evita movimiento de pantalla
objExcel.Application.ScreenUpdating = True
'Parùmetro evitar mostrar pop ups de Excel
objExcel.Application.DisplayAlerts = False

'Abre libro Excel
Set objWorkbookNomina = objExcel.Workbooks.Open(WorkbookPathNomina)
Set objWorksheetNomina = objWorkbookNomina.Worksheets(WorkbookSheetNomina)

Set objWorkbookReclas = objExcel.Workbooks.Open(WorkbookPathReclas)
Set objWorksheetReclas = objWorkbookReclas.Worksheets(WorkbookSheetReclas)

Set objWorkbookCatalogo = objExcel.Workbooks.Open(WorkbookPathCatalogo)
Set objWorksheetCatalogo = objWorkbookCatalogo.Worksheets(WorkbookSheetCatalogo)

'Obtener ultima fila usando cmo base la columna C
ultimaFila = objWorksheetReclas.Cells(objWorksheetReclas.Rows.Count,3).End(-4162).Row

ultimaFilaC = objWorksheetCatalogo.Cells(objWorksheetCatalogo.Rows.Count,1).End(-4162).Row

'Obtener el Centro de Costo y el valor del bloque MX29
For i = 1 to ultimaFila
    Set rng = objWorksheetReclas.Range("A" & i)
   'Se itera validando el valor de la celda
    If rng.value <> "" And rng.value <> "Row Labels" And rng.value <> "Grand Total" Then
        If Not rng.PivotTable Is Nothing Then
            'Se establece el rango de la Pivot
            Set pf = rng.PivotField
            'Evalua si es el Centro de costo o el bloque MX29 o MX09
            If pf.Name = "Sender Cost Center" Then
                bloque = rng.Value
            ElseIf pf.Name = "Object" And (InStr(CStr(rng.value),"MX29") <> 0 Or InStr(CStr(rng.value),"MX09") <> 0) Then
                'Se agregan los valores a un diccionario
                objDictionary.Add rng.Value, bloque
                objDictionaryPorcentual.Add rng.Value, objWorksheetReclas.Range("C" & i)
            End If
        End If
    End If
    Set rng = Nothing
Next
Set pf = Nothing

'Se cargan los valores de llaves y elementos
d_key = objDictionary.Keys
d_item = objDictionary.Items
p_key = objDictionaryPorcentual.Keys
p_item = objDictionaryPorcentual.Items

'Itera sobre el valor de la llave y el elemento
'Valida sobre el valor del Catalogo
For i = 0 To objDictionary.Count -1 
    For j = 1 To ultimaFilaC
        If d_item(i) = objWorksheetCatalogo.Cells(j,1).value Then
            d_item(i) = objWorksheetCatalogo.Cells(j,2).value
        End If
    Next
Next

objWorkbookCatalogo.Save
objWorkbookCatalogo.Close SaveChanges = True

Const xlPart = 2
Const xlValues = -4163

'Obtener ultima fila usando cmo base la columna C
lastRow = objWorksheetNomina.Cells(objWorksheetNomina.Rows.Count,8).End(-4162).Row
lastCol = objWorksheetNomina.Cells(3, objWorksheetNomina.Columns.Count).End(-4159).Column - 1 ' -4159 = xlToLeft


For i = 4 To lastRow
    If objWorksheetNomina.Cells(i, 8).Value = "" Then
        lastRow = i + 1
        Exit For
    End If
Next

'Establece rango de busqueda
Set mRange = objWorksheetNomina.Range(objWorksheetNomina.Cells(1,1),objWorksheetNomina.Cells(10,lastCol))
Set iRange = objWorksheetNomina.Range("H:H")
Dim iFind, kFind

'Busca dependioendo del valor de la llave y el elemento
'para obtener fila, columna y establecer el valor porcentual
For i = 0 To objDictionary.Count -1 
    'Busca el Nombre para obtener la fila
    Set iFind = iRange.Find(d_item(i),,xlValues,xlPart)
    If iFind Is Nothing Then
        ' Insertar una nueva fila en la posiciùn lastRow
        objWorksheetNomina.Rows(lastRow).Insert -4121 ' -4121 = xlDown
        ccRow = lastRow
    Else
        ccRow = iFind.Row
    End IF

    Set kFind = mRange.Find(d_key(i),,xlValues,xlPart)
    If kFind Is Nothing Then
        ' Insertar una nueva columna en la posiciùn 3 (por ejemplo, antes de la columna C)
        objWorksheetNomina.Columns(lastCol - 1).Copy
        objWorksheetNomina.Columns(lastCol).Insert -4121 ' -4121 = xlToRight
        ' Quitar el modo de corte/copia
        objExcel.CutCopyMode = False
        objWorksheetNomina.Cells(3, lastCol).Value = d_key(i)
        ccCol = lastCol
    Else
        ccCol = kFind.Column
    End IF

    objWorksheetNomina.Cells(ccRow, ccCol).Value = p_item(i)

    Set iFind = Nothing
    Set kFind = Nothing
Next

Set objDictionary = Nothing
Set objDictionaryPorcentual = Nothing

'Guarda y cierre el libro
objWorkbookReclas.Save
objWorkbookReclas.Close SaveChanges = True

objWorkbookNomina.Save
objWorkbookNomina.Close SaveChanges = True

'Quita la instancia del objeto Excel
objExcel.Quit

'Devuelve el error en caso de
If Err.Number <> 0 Then
    Msg = "Error was generated by " & Err.Source & ". " & Err.Description
    WScript.Echo Msg
End if