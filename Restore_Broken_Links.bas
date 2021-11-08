Attribute VB_Name = "Restore_Broken_Links"
Option Explicit

Sub Save_Files()

	'Run this!

    'Solution from here:
    'https://stackoverflow.com/q/41068315
	'based on the code here:
	'https://stackoverflow.com/a/62667982
    
    Dim dialogBox As FileDialog
    Dim sourceFullName As String
    Dim sourceFilePath As String
    Dim sourceFileName As String
    Dim sourceFileType As String
    Dim newFileName As Variant
    Dim tempFileName As String
    Dim zipFilePath As Variant
    Dim oApp As Object
    Dim fso As Object
    Dim xmlSheetFile As String
    Dim xmlFile As Integer
    Dim xmlFileContent As String
    Dim xmlStartProtectionCode As Double
    Dim xmlEndProtectionCode As Double
    Dim xmlProtectionString As String
    
    Dim thing As Variant
    Dim fldr As Variant
    Dim fls As Variant
    Dim newFolderPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Open dialog box to select a file
    Set dialogBox = Application.FileDialog(msoFileDialogFolderPicker)
    dialogBox.AllowMultiSelect = False
    dialogBox.Title = "Select a folder where the Files are located"
    
    If dialogBox.Show = -1 Then
        sourceFullName = dialogBox.SelectedItems(1)
    Else
        Exit Sub
    End If

    Set fldr = fso.GetFolder(sourceFullName)
    Set fls = fldr.Files

    For Each thing In fls
        Save_Files_Replacer (sourceFullName & "\" & thing.Name)
    Next
    
    'Show message box
    MsgBox "The fixed Excel files have been saved to the folder ""Output"" in the same directory." & vbCrLf & "Please move them and delete the working folder afterwards.", vbInformation + vbOKOnly, Title:="Done"

End Sub

Private Sub Save_Files_Replacer(inputtext As String)

    'File Handling Variables
    Dim dialogBox As FileDialog
    Dim sourceFullName As String
    Dim sourceFilePath As String
    Dim sourceFileName As String
    Dim sourceFileType As String
    Dim newFileName As Variant
    Dim tempFileName As String
    Dim zipFilePath As Variant
    Dim tempFilePath As String
    Dim oApp As Object
    Dim fso As Object
    Dim xmlSheetFile As String
    Dim xmlFile As Integer
    Dim xmlFileContent As String
    Dim xmlStartProtectionCode As Double
    Dim xmlEndProtectionCode As Double
    Dim xmlProtectionString As String
    
    'Regex Variables
    Dim regex As Object
    Dim str As String
    Dim matches As Variant
    Dim Match As Variant
    Dim subMatch As Variant
    Dim int_i As Long
    Dim replacestr As String
    Set regex = CreateObject("VBScript.RegExp")
    
    sourceFullName = inputtext
    
    'Get folder path, file type and file name from the sourceFullName
    sourceFilePath = Left(sourceFullName, InStrRev(sourceFullName, "\"))
    sourceFileType = Mid(sourceFullName, InStrRev(sourceFullName, ".") + 1)
    sourceFileName = Mid(sourceFullName, Len(sourceFilePath) + 1)
    sourceFileName = Left(sourceFileName, InStrRev(sourceFileName, ".") - 1)
    
    'If the file is a temporary one, we don't do all of this
    If Left(sourceFileName, 1) = "~" Then
        Exit Sub
    End If
    
    'Use the date and time to create a unique file name
    tempFileName = "Temp" & Format(Now, "_yyyymmdd_hhmmss")
    tempFilePath = Environ("TEMP") & "\"
    
    'Copy and rename original file to a zip file with a unique name
    newFileName = tempFilePath & tempFileName & ".zip"
    On Error Resume Next
    FileCopy sourceFullName, newFileName
    
    If Err.Number <> 0 Then
        MsgBox "Unable to copy " & sourceFullName & vbNewLine & "Check the file is closed and try again"
        Exit Sub
    End If
    On Error GoTo 0
    
    'Create folder to unzip to
    zipFilePath = tempFilePath & tempFileName & "\"
    MkDir zipFilePath
    
    'Extract the files into the newly created folder
    Set oApp = CreateObject("Shell.Application")
    oApp.Namespace(zipFilePath).CopyHere oApp.Namespace(newFileName).items
    
    'loop through each file in the \xl\worksheets folder of the unzipped file
    xmlSheetFile = Dir(zipFilePath & "\xl\externalLinks\_rels\*.rels")
    Do While xmlSheetFile <> ""
    
        'Read text of the file to a variable
        xmlFile = FreeFile
        Open zipFilePath & "xl\externalLinks\_rels\" & xmlSheetFile For Input As xmlFile
        xmlFileContent = Input(LOF(xmlFile), xmlFile)
        Close xmlFile
        
        'All of this is roughly equivalent to the following Regex:
        ' 's/(\<Relationship\ Id\=\"rId2\".*?\/\>)(\<Relationship\ Id\=\"rId1\".*?Target=")(.*[\\|\/])?(.*?xls.)(.*?)?(\".*?\/\>)/$2$4$6/'
    
        'Replace the text in files by using Regex
        'https://analystcave.com/excel-regex-tutorial/#Regex_Replace_pattern_in_a_string
        With regex
            'VBA needs double quotes around a single " in a Regex
            .Pattern = "(\<Relationship\ Id\=\""rId2\"".*?\/\>)(\<Relationship\ Id\=\""rId1\"".*?Target="")(.*[\\|\/])?(.*?xls.)(.*?)?(\"".*?\/\>)"
            .Global = True
        End With
        
        'Setup the text for the regex
        str = xmlFileContent
        replacestr = ""
         
        'Search for matches
        Set matches = regex.Execute(str)
          
        'Output the matches
        For Each Match In matches
            If Match.SubMatches.Count > 0 Then
                For int_i = 0 To Match.SubMatches.Count - 1
                    subMatch = Match.SubMatches(int_i)
                    
                    'Print groups like this: $2$4$6
                    If int_i = 1 Or int_i = 3 Or int_i = 5 Then
                        replacestr = replacestr & subMatch
                    End If
                Next int_i
            End If
        Next Match
        
        'Replace the match with our output matched groups
        xmlFileContent = regex.Replace(str, replacestr)
    
        'Output the text of the variable to the file
        xmlFile = FreeFile
        Open zipFilePath & "xl\externalLinks\_rels\" & xmlSheetFile For Output As xmlFile
        Print #xmlFile, xmlFileContent
        Close xmlFile
    
        'Loop to next xmlFile in directory
        xmlSheetFile = Dir
    
    Loop
    
    'Create empty Zip File
    Open tempFilePath & tempFileName & ".zip" For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
    
    'Move files into the zip file
    oApp.Namespace(tempFilePath & tempFileName & ".zip").CopyHere oApp.Namespace(zipFilePath).items
    'Keep script waiting until Compressing is done
    On Error Resume Next
    Do Until oApp.Namespace(tempFilePath & tempFileName & ".zip").items.Count = oApp.Namespace(zipFilePath).items.Count
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop
    On Error GoTo 0
    
    'Delete the files & folders created during the sub
    Set fso = CreateObject("scripting.filesystemobject")
    fso.deletefolder tempFilePath & tempFileName
    
    On Error Resume Next
    fso.CreateFolder (sourceFilePath & "Output")
    On Error GoTo 0
    
    'Rename the final file back to an xlsx file
    Name tempFilePath & tempFileName & ".zip" As sourceFilePath & "Output\" & sourceFileName & "." & sourceFileType

End Sub

