Sub Main
	Call ExcelImport()	'C:\Users\Intel\Documents\Mis documentos IDEA\Samples\Archivos fuente.ILB\Ejemplo.xlsx
End Sub


' Archivo - Asistente de importación: Excel
Function ExcelImport
	Set task = Client.GetImportTask("ImportExcel")
	dbName = "C:\Users\Intel\Documents\Mis documentos IDEA\Samples\Archivos fuente.ILB\Ejemplo.xlsx"
	task.FileToImport = dbName
	task.SheetToImport = "Inventario"
	task.OutputFilePrefix = "Ejemplo"
	task.FirstRowIsFieldName = "TRUE"
	task.EmptyNumericFieldAsZero = "TRUE"
	task.PerformTask
	dbName = task.OutputFilePath("Inventario")
	Set task = Nothing
	Client.OpenDatabase(dbName)
End Function