Sub CreateInvertVarTable()
        'приложение Inventor
    Dim oApp As Inventor.Application
    
    ' Получим ссылку на активное приложение INVENTOR
    Set oApp = ThisApplication
    
    ' Получим ссылку на активный документ.
    Dim oDrawDoc As DrawingDocument
    Set oDrawDoc = oApp.ActiveDocument


    ' Get the PropertySets object.
    Set oPropSets = oDrawDoc.PropertySets

    ' Get the design tracking property set.
    Set oPropSet = oPropSets.Item("Design Tracking Properties")
    
    Set oPartNumiProp = oPropSet.Item("Part Number")
    modelMarking = oPartNumiProp.Value

    ' Set a reference to the active sheet.
    Dim oSheet As Sheet
    Set oSheet = oDrawDoc.ActiveSheet
    
    ' Set the column titles
    Dim oTitles(0 To 3) As String
    oTitles(0) = "№"
    oTitles(1) = "Обозначение"
    oTitles(2) = "Материал"
    oTitles(3) = "Покрытие"
    
    ' Set the contents of the custom table (contents are set row-wise)
    Dim oContents(0 To 7) As String
    oContents(0) = "1"
    oContents(1) = modelMarking
    'oContents(1) = "<Property Document='Drawing' FormatID='{32853F0F-3444-11d1-9E93-0060B03C1CA6}' PropertyID='5' DispName='ОБОЗНАЧЕНИЕ' Precision='-1' />"
    oContents(2) = "Лист Б-ПН-3 ГОСТ 19903-74/ Ст3сп ГОСТ 14637-89"
    oContents(3) = "Покрытие: III; У1; в соответствии с заказом."
    
    oContents(4) = "2"
    oContents(5) = modelMarking + "-01"
    'oContents(5) = "<Property Document='Drawing' FormatID='{32853F0F-3444-11d1-9E93-0060B03C1CA6}' PropertyID='5' DispName='ОБОЗНАЧЕНИЕ' Precision='-1' />"
    oContents(6) = "Лист ОЦ  Б-ПН-3,0 ГОСТ 14918-80/ Ст08кп ГОСТ 1050-88"
    oContents(7) = "-"
        
    ' Set the column widths (defaults to the column title width if not specified)
    Dim oColumnWidths(0 To 3) As Double
    oColumnWidths(0) = 0.5
    oColumnWidths(1) = 6.3
    oColumnWidths(2) = 8.08
    oColumnWidths(3) = 6
    
    Dim oRowHeights(0 To 1) As Double
    oRowHeights(0) = 1.05
    oRowHeights(1) = 1.05
    
    Dim oPoint As Point2d
    Set oPoint = ThisApplication.TransientGeometry.CreatePoint2d(2, 3.14)
    
    ' Create the custom table
    Dim oCustomTable As CustomTable
    Set oCustomTable = oSheet.CustomTables.Add("Таблица исполнений", oPoint, 4, 2, oTitles, oContents, oColumnWidths, oRowHeights)
                                        
  
    ' Create a table format object
    Dim oFormat As TableFormat
    Set oFormat = oSheet.CustomTables.CreateTableFormat
    
    ' Set inside line color to red.
    'oFormat.InsideLineColor = ThisApplication.TransientObjects.CreateColor(255, 0, 0)
    
    ' Set outside line weight.
    oFormat.OutsideLineWeight = 0.05
    
      ' Set outside line weight.
    oFormat.InsideLineWeight = 0.01
    
    ' Modify the table formats
    oCustomTable.OverrideFormat = oFormat
    
    'MsgBox oCustomTable
    ' Change the 3rd column to be left justified.
    oCustomTable.Columns.Item(1).ValueHorizontalJustification = kAlignTextCenter
    oCustomTable.Columns.Item(4).ValueHorizontalJustification = kAlignTextCenter
End Sub
