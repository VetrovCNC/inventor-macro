Sub ChangeReferencedDocument()

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
    
    'Dim oRefDocs As Document.ReferencedDocuments
    'Set oRefDocs = oDoc.ReferencedDocuments()
    'MsgBox getFileName(Split(oDoc.FullDocumentName, ".idw")(0))
    'MsgBox oDoc.File.ReferencedFileDescriptors.Count

    Dim oRefFileDesc As ReferencedFileDescriptor
    For Each oRefFileDesc In oDoc.ReferencedFileDescriptors
        'MsgBox oRefFileDesc.FullFileName
        MsgBox getFileExt(oRefFileDesc.FullFileName)
    Next
    
    ' Create a new FileDialog object.
    Dim oFileDlg As FileDialog
    Call ThisApplication.CreateFileDialog(oFileDlg)

    ' Define the filter to select part and assembly files or any file.
    oFileDlg.Filter = "Inventor Files (*.iam;*.ipt)|*.iam;*.ipt|All Files (*.*)|*.*"

    ' Define the part and assembly files filter to be the default filter.
    oFileDlg.FilterIndex = 1

    ' Set the title for the dialog.
    oFileDlg.DialogTitle = "Open File Test"

    ' Set the initial directory that will be displayed in the dialog.
    oFileDlg.InitialDirectory = getPathName(oDoc.FullFileName)

    ' Set the flag so an error will be raised if the user clicks the Cancel button.
    oFileDlg.CancelError = True

    ' Show the open dialog.  The same procedure is also used for the Save dialog.
    ' The commented code can be used for the Save dialog.
    On Error Resume Next
    oFileDlg.ShowOpen
    'oFileDlg.ShowSave

    ' If an error was raised, the user clicked cancel, otherwise display the filename.
    If Err Then
        'MsgBox "User cancelled out of dialog"
    ElseIf oFileDlg.FileName <> "" Then
        'MsgBox "File " & oFileDlg.FileName & " was selected."
        Call oDoc.File.ReferencedFileDescriptors.Item(1).ReplaceReference(oFileDlg.FileName)
        Call oDoc.Update
        Call UpdateMarking
    End If
    
End Sub