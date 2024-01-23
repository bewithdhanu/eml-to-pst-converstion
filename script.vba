Sub CreatePSTAndAddEML()
    ' Create Outlook Application object
    Dim outlookApp As Object
    Set outlookApp = CreateObject("Outlook.Application")
    
    ' Create MAPI Namespace
    Dim outlookNamespace As Object
    Set outlookNamespace = outlookApp.GetNamespace("MAPI")
    
    ' Create a new PST file
    Dim pstFolder As Object
    Set pstFolder = outlookNamespace.Folders.Add("NewPST", olFolderOutlook)
    pstFolder.DefaultItemType = olMailItem
    pstFolder.Display
    
    ' Import EML file into the new PST
    Dim emlFilePath As String
    emlFilePath = "C:\Path\To\Your\File.eml" ' Replace with the actual path to your EML file
    
    If Dir(emlFilePath) <> "" Then
        ' Import the EML file
        pstFolder.Import emlFilePath, olImportItemEML
    Else
        MsgBox "EML file not found at: " & emlFilePath, vbExclamation
    End If
End Sub
