Public Sub AddMassToIPart()
	' Макрос работает в контексте модели детали. 
	' Деталь должна быть создана параметрической с последним столбцом под массу
	' Этот параметр содается на вкладке Прочее диалогового окна Свойства Inventor
	Dim oDoc As PartDocument
	Set oDoc = ThisApplication.ActiveDocument
	oDoc.UnitsOfMeasure.MassUnits = kKilogramMassUnits
	Dim oFactory As iPartFactory
	Set oFactory = oDoc.ComponentDefinition.iPartFactory
	' чтобы редактировать точность измерения массы, измените цифру 3 (количество знаков после запятой) на нужное вам значение
	' если нужно убрать единицы измерения, сотрите & "кг" в последней строке
	For i = 1 To oFactory.TableRows.Count
		oFactory.DefaultRow = oFactory.TableRows(i)
		oDoc.Update2 (True)
		oFactory.TableRows(i)(oFactory.TableColumns.Count).Value = Math.Round(oDoc.ComponentDefinition.MassProperties.Mass, 3) & " кг"
	Next
End Sub