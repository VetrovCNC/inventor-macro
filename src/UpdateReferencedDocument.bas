Function GetFilenameFromPath(ByVal strPath As String) As String
' Returns the rightmost characters of a string upto but not including the rightmost '\'
' e.g. 'c:\winnt\win.ini' returns 'win.ini'

    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function

Function getFileExt(pf) As String: getFileExt = Mid(pf, InStrRev(pf, ".")): End Function

Function getFileExtV2(pf) As String: getFileExtV2 = Mid(pf, InStrRev(pf, "-")): End Function

Function getFileName(pf) As String: getFileName = Mid(pf, InStrRev(pf, "\") + 1): End Function

Function getPathName(pf) As String: getPathName = Left(pf, InStrRev(pf, "\")): End Function

Sub UpdateReferencedDocument()

    'приложение Inventor
    Dim oApp As Inventor.Application
    
    ' Получим ссылку на активное приложение INVENTOR
    Set oApp = ThisApplication
    
    ' Получим ссылку на активный документ.
    Dim oDoc As Document
    Set oDoc = oApp.ActiveDocument
    
    'Проверка: а в чертеже ли мы?
    If oDoc.DocumentType <> kDrawingDocumentObject Then
        MsgBox "Процедура предназначена для работы в контексте чертежа."
        Exit Sub
    End If

    Dim count As Long
    count = 1
    
    Dim file_str As String
    file_str = getPathName(oDoc.FullFileName) + getFileName(Split(oDoc.FullDocumentName, ".idw")(0))

    Dim file_name As String
    
    Dim oRefFileDesc As ReferencedFileDescriptor
    For Each oRefFileDesc In oDoc.ReferencedFileDescriptors
        'MsgBox oRefFileDesc.FullFileName
        'MsgBox count
        If count = 1 Then
            file_name = file_str + getFileExt(oRefFileDesc.FullFileName)
        Else
            file_name = file_str + getFileExtV2(oRefFileDesc.FullFileName)
        End If
        Call oDoc.File.ReferencedFileDescriptors.Item(count).ReplaceReference(file_name)
        count = count + 1
    Next
    
    Call oDoc.Update

End Sub
