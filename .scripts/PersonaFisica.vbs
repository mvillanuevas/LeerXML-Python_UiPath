On Error Resume Next

Set objArgs = WScript.Arguments

WorkbookPathData = objArgs(0)
WorkbookSheet = objArgs(1)
UUID = objArgs(2)
Fecha = objArgs(3)
Monto = objArgs(4)
MovimientoISR = objArgs(5)
Nombre = objArgs(6)

'Genera un objeto de tipo Excel Application
Set objExcel = CreateObject("Excel.Application")

'Parámetro para indicar si se quiere visible la aplicación de Excel
objExcel.Application.Visible = False
'Evita movimiento de pantalla
objExcel.Application.ScreenUpdating = False
'Parámetro evitar mostrar pop ups de Excel
objExcel.Application.DisplayAlerts = False

'Abre libro Excel
Set objWorkbook = objExcel.Workbooks.Open(WorkbookPathData)
Set objWorksheet = objWorkbook.Worksheets(WorkbookSheet)

Const xlPart = 2
Const xlValues = -4163

'Establece rango de busqueda
Set mRange = objWorksheet.Range("A:A")

'Busca el Nombre para obtener la fila
Dim mFind :	Set mFind = mRange.Find(Nombre,,xlValues,xlPart)

'Inserta valortes de XML
If Not mFind Is Nothing Then
	firstAddress = mFind.Address
	Do
		If objWorksheet.Cells(mFind.Row, 3).value = "" Then
			objWorksheet.Cells(mFind.Row, 3).value = UUID
			objWorksheet.Cells(mFind.Row, 4).value = Fecha
			objWorksheet.Cells(mFind.Row, 5).value = Monto
			objWorksheet.Cells(mFind.Row, 6).value = MovimientoISR
		End If
		Set mFind =  mRange.FindNext(mFind)
	Loop While Not mFind Is Nothing And mFind.Address <> firstAddress
End If

'Guarda y cierre el libro
objWorkbook.Save
objWorkbook.Close SaveChanges = True

'Quita la instancia del objeto Excel
objExcel.Quit

'Devuelve el error en caso de
If Err.Number <> 0 Then
    Msg = "Error was generated by " & Err.Source & ". " & Err.Description
    WScript.Echo Msg
End if