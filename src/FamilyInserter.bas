Sub FamilyInserter()

    'Фамилии в штампе чертежа
    Dim modelRazrabotal As String
    Dim modelProveril As String
    Dim modelTkontr As String
    Dim modelNormokontrol As String
    Dim modelUtverdil As String
    Dim modelNachalnikKB As String

    'Разработал
    modelRazrabotal = ""
        
    'Проверил
    modelProveril = ""
    
    'Т. контр
    modelTkontr = ""
    
    'Начальник КБ
    modelNachalnikKB = ""
    
    'Нормоконтроль
    modelNormokontrol = ""
    
    'Утвердил
    modelUtverdil = ""
    

    'приложение Inventor
    Dim oApp        As Inventor.Application
    
    ' Получим ссылку на активное приложение INVENTOR
    Set oApp = ThisApplication
   
    ' Get the active document.
    Dim oDoc As Document
    Set oDoc = oApp.ActiveDocument

    ' Get the PropertySets object.
    Dim oPropSets As PropertySets
    Set oPropSets = oDoc.PropertySets

    ' Get the design tracking property set.
    Dim oPropSet As PropertySet
    Set oPropSet = oPropSets.Item("Design Tracking Properties")

    ' Get the part number iProperty.
    Dim oPartNumiProp As Property
    
    ' Get the part number iProperty.
    Dim oPartDescrProp As Property
    Set oPartDescrProp = oPropSet.Item("Description")
    
    'MsgBox "The part number is: " & oPartDescrProp.Value
    
    'Инициализация транзакции для возможной отмены за один шаг
    Dim oConstrTransaction As Transaction
    Set oConstrTransaction = oApp.TransactionManager. _
             StartTransaction(oDoc, "Заполнение фамилий")
    
    'Разработал (вкладка Проект)
    Set oPartNumiProp = oPropSet.Item("Designer")
    oPartNumiProp.Value = modelRazrabotal
    
    'Проверил (вкладка Статус)
    Set oPartNumiProp = oPropSet.Item("Checked By")
    oPartNumiProp.Value = modelProveril
    
    'Нач. отдела (вкладка Проект)
    Set oPartNumiProp = oPropSet.Item("Authority")
    oPartNumiProp.Value = modelNachalnikKB
    
    'Нормоконтроль (вкладка Статус)
    Set oPartNumiProp = oPropSet.Item("Engr Approved By")
    oPartNumiProp.Value = modelNormokontrol
    
    'Т.контр (вкладка Проект)
    Set oPartNumiProp = oPropSet.Item("Engineer")
    oPartNumiProp.Value = modelTkontr
    
    'Утвердил (вкладка Статус)
    Set oPartNumiProp = oPropSet.Item("Mfg Approved By")
    oPartNumiProp.Value = modelUtverdil
    
    ' Get the Inventor Summary Information.
    Set oPropSet = oPropSets.Item("Inventor Summary Information")
    
    'Разраб. (вкладка Документ)
    Set oPartNumiProp = oPropSet.Item("Author")
    oPartNumiProp.Value = modelRazrabotal
    
    'Наименование (вкладка Документ)
    Set oPartNumiProp = oPropSet.Item("Title")
    oPartNumiProp.Value = oPartDescrProp.Value

    ' Get the Inventor Document Summary Information.
    Set oPropSet = oPropSets.Item("Inventor Document Summary Information")
    
    'Нач. отд. (вкладка Документ)
    Set oPartNumiProp = oPropSet.Item("Manager")
    oPartNumiProp.Value = modelNachalnikKB
    
    'Организация (вкладка Документ)
    Set oPartNumiProp = oPropSet.Item("Company")
    oPartNumiProp.Value = "ООО ""ЗЭО"""

oConstrTransaction.End 'завершение транзакции

End Sub