Sub UpdateMarking()

    'приложение Inventor
    Dim oApp As Inventor.Application
    
    ' Получим ссылку на активное приложение INVENTOR
    Set oApp = ThisApplication
    
    ' Получим ссылку на активный документ.
    Dim oDoc As Document
    Set oDoc = oApp.ActiveDocument
    Dim sFN As String
    sFN = CreateObject("Scripting.FileSystemObject").GetBaseName(oDoc.FullFileName())
    
    ' Get the PropertySets object.
    Dim oPropSets As PropertySets
    Set oPropSets = oDoc.PropertySets

    ' Get the design tracking property set.
    Dim oPropSet As PropertySet
    Set oPropSet = oPropSets.Item("Design Tracking Properties")
    
    ' Get the part number iProperty.
    Dim oPartNumiProp As Property
    Set oPartNumiProp = oPropSet.Item("Part Number")
    oPartNumiProp.Value = sFN
    
    'Проверка: а в модели ли мы?
    If oDoc.DocumentType = kAssemblyDocumentObject Or oDoc.DocumentType = kPartDocumentObject Then
        'MsgBox "Это модель."
        Dim partName As String
        Set oPropSet = oPropSets.Item("Design Tracking Properties")
        Set oPartNumiProp = oPropSet.Item("Part Number")
        oPartNumiProp.Value = sFN
        Set oPartNumiProp = oPropSet.Item("Description")
        partName = oPartNumiProp.Value
        
        Set oPropSet = oPropSets.Item("Inventor Summary Information")
    
        'Наименование (вкладка Документ)
        Set oPartNumiProp = oPropSet.Item("Title")
        oPartNumiProp.Value = partName
        
    Else
        'MsgBox "Это чертеж."
        Dim fName As String
        fName = oDoc.File.ReferencedFileDescriptors.Item(1).FullFileName
        'MsgBox fName
        
        Dim oRefDoc As Document
        Set oRefDoc = oApp.Documents.Open(fName)
        
        Dim modelName As String
        Dim modelMarking As String
        
        'Фамилии в штампе чертежа
        Dim modelRazrabotal As String
        Dim modelProveril As String
        Dim modelTkontr As String
        Dim modelNormokontrol As String
        Dim modelUtverdil As String
        Dim modelNachalnikKB As String
        
        ' Get the PropertySets object.
        Set oPropSets = oRefDoc.PropertySets
    
        ' Get the design tracking property set.
        Set oPropSet = oPropSets.Item("Design Tracking Properties")
        
        Set oPartNumiProp = oPropSet.Item("Part Number")
        modelMarking = oPartNumiProp.Value
        
        Set oPartNumiProp = oPropSet.Item("Description")
        modelName = oPartNumiProp.Value
        
        'Разработал (вкладка Проект)
        Set oPartNumiProp = oPropSet.Item("Designer")
        modelRazrabotal = oPartNumiProp.Value
        
        'Проверил (вкладка Статус)
        Set oPartNumiProp = oPropSet.Item("Checked By")
        modelProveril = oPartNumiProp.Value
        
        'Нач. отдела (вкладка Проект)
        Set oPartNumiProp = oPropSet.Item("Authority")
        modelNachalnikKB = oPartNumiProp.Value
        
        'Нормоконтроль (вкладка Статус)
        Set oPartNumiProp = oPropSet.Item("Engr Approved By")
        modelNormokontrol = oPartNumiProp.Value
        
        'Т.контр (вкладка Проект)
        Set oPartNumiProp = oPropSet.Item("Engineer")
        modelTkontr = oPartNumiProp.Value
        
        'Утвердил (вкладка Статус)
        Set oPartNumiProp = oPropSet.Item("Mfg Approved By")
        modelUtverdil = oPartNumiProp.Value
        
        'Call oRefDoc.Close
        
        Call oDoc.Activate
        
        ' Get the PropertySets object.
        Set oPropSets = oDoc.PropertySets
        
        ' Get the design tracking property set.
        Set oPropSet = oPropSets.Item("Design Tracking Properties")
        
        Set oPartNumiProp = oPropSet.Item("Part Number")
        oPartNumiProp.Value = modelMarking
        
        Set oPartNumiProp = oPropSet.Item("Description")
        oPartNumiProp.Value = modelName
        
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
        
        Set oPropSet = oPropSets.Item("Inventor Summary Information")
    
        'Наименование (вкладка Документ)
        Set oPartNumiProp = oPropSet.Item("Title")
        oPartNumiProp.Value = modelName
        
        'Разраб. (вкладка Документ)
        Set oPartNumiProp = oPropSet.Item("Author")
        oPartNumiProp.Value = modelRazrabotal
        
        ' Get the Inventor Document Summary Information.
        Set oPropSet = oPropSets.Item("Inventor Document Summary Information")
        
        'Нач. отд. (вкладка Документ)
        Set oPartNumiProp = oPropSet.Item("Manager")
        oPartNumiProp.Value = modelNachalnikKB
        
    End If
    
End Sub