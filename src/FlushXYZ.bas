'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                 Flush_XYZ
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Выделенный пользователем компонент фиксируется
'в координатах сборки наложением трех зависимостей
'совмещения заподлицо одноименных базовых
'плоскостей XY, YZ и XZ сборки и компонента
'
'Процедура работает в контексте сборки.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 
Public Sub Flush_XYZ()
 
   Dim oApp        As Inventor.Application        'приложение Inventor
   Dim oAsmCompDef As AssemblyComponentDefinition 'сборка
   Dim oCompOcc    As ComponentOccurrence         'компонент
   Dim oSelectSet  As SelectSet
 
   Dim oAsmPlane       As WorkPlane      'рабочие плоскости сборки
   Dim oPartPlane      As WorkPlane      'рабочие плоскости детали
   Dim oPartPlaneProxy As WorkPlaneProxy 'proxy-плоскости детали
 
   Dim i As Long  'счетчик плоскостей 1,2,3
 
   ' Получим ссылку на активное приложение INVENTOR
   Set oApp = ThisApplication
 
   'Проверка: а в сборке ли мы?
   If oApp.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
       MsgBox "Процедура предназначена для работы в контексте сборки."
       Exit Sub
   End If
   
   'Получим ссылку на определение компонентов активной сборки
   Set oAsmCompDef = oApp.ActiveDocument.ComponentDefinition
 
   'Ссылка на коллекцию SelectSet активного документа
   Set oSelectSet = oApp.ActiveDocument.SelectSet
 
   'Проверка: должен быть выделен один и только один элемент
   If oSelectSet.Count <> 1 Then
       MsgBox "Следует выделить один компонент."
       Exit Sub
   End If
 
   'Проверка: должен быть выделен именно компонент
   If Not (TypeOf oSelectSet.Item(1) Is ComponentOccurrence) Then
       MsgBox "Следует выделить один компонент."
       Exit Sub
   End If
 
   'ссылка на выделенный компонент (деталь)
   Set oCompOcc = oSelectSet.Item(1)
 
   'Инициализация транзакции для возможной отмены за один шаг
   Dim oConstrTransaction As Transaction
   Set oConstrTransaction = oApp.TransactionManager. _
            StartTransaction(oApp.ActiveDocument, "Привязка_XYZ")
 
 
   'Совмещаем в цикле базовые плоскости сборки и компонента.
 
   ' i=1:   плоскость YZ,   нормаль - ось X
   ' i=2:   плоскость ZX,   нормаль - ось Y
   ' i=3:   плоскость XY,   нормаль - ось Z
 
   For i = 1 To 3
 
      'ссылка на базовую плоскость i сборки
      Set oAsmPlane = oAsmCompDef.WorkPlanes.Item(i)
 
      'ссылка на базовую плоскость i выделенного компонента
      Set oPartPlane = oCompOcc.Definition.WorkPlanes.Item(i)
 
      'создаем прокси-объект рабочей плоскости компонента
      Call oCompOcc.CreateGeometryProxy(oPartPlane, oPartPlaneProxy)
 
      'создаем зависимость совмещения рабочих плоскостей типа "Заподлицо"
      'плоскости сборки и прокси-объекта плоскости компонента
      Call oAsmCompDef.Constraints.AddFlushConstraint(oAsmPlane, oPartPlaneProxy, 0)
 
   Next i
 
   oConstrTransaction.End 'завершение транзакции
 
End Sub   ' Flush_XYZ
