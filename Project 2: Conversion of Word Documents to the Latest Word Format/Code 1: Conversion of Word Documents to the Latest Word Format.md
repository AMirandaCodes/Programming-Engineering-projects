# Code 1: Conversion of Word Documents to the Latest Word Format

## Code Description
This VBA script automates the process of converting all Word documents in a selected folder from the `.doc` format to the latest `.docx` format. The original `.doc` files are replaced by the newly converted `.docx` versions.

## Code Features
- **Folder Selection**: Prompts the user to select a folder containing the Word documents for conversion.
- **Same Destination**: The converted `.docx` files are saved in the same folder as the original `.doc` files.
- **Automatic Replacement**: Deletes the original `.doc` files after successful conversion, ensuring no duplicates.

![image](https://github.com/user-attachments/assets/227efdc5-724a-4e6f-b40b-58ffca445dd7)

## Code Usage Instructions
1. Open Microsoft Word.
2. Access the VBA editor (Alt + F11) and paste the script into a module.
3. Run the script from the VBA editor.
4. Select the folder containing the `.doc` files when prompted.
5. The converted `.docx` files will replace the originals in the same folder.

## Limitations
- **One-Off Execution**: Designed for single-use conversions. For repeated or scheduled tasks, consider automation enhancements.
- **Layout Adjustments**: Minor layout changes may occur during the conversion process, especially with older documents containing complex formatting.
- **File Deletion**: The script deletes original `.doc` files after conversion. Ensure you have backups if required.

## Code Snippet
```vba
Sub ConvertToLatestWordFormat()
    Dim strFolderPath As String
    Dim strFileName As String
    Dim doc As Document
    Dim convertedFileName As String
    
    ' Prompt user to select a folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing Word Documents"
        If .Show = -1 Then
            strFolderPath = .SelectedItems(1) & "\"
        Else
            MsgBox "No folder selected. Process cancelled."
            Exit Sub
        End If
    End With

    ' Get the first .doc file in the folder
    strFileName = Dir(strFolderPath & "*.doc")
    
    ' Loop through all .doc files in the folder
    Do While strFileName <> ""
        ' Open the document
        Set doc = Documents.Open(strFolderPath & strFileName, ReadOnly:=False)
        
        ' Set the new file name (change extension to .docx)
        convertedFileName = Left(strFileName, InStrRev(strFileName, ".")) & "docx"
        
        ' Save the document in the latest Word format (.docx)
        doc.SaveAs2 FileName:=strFolderPath & convertedFileName, FileFormat:=wdFormatXMLDocument
        
        ' Close the document
        doc.Close SaveChanges:=False
        
        ' Delete the original .doc file
        Kill strFolderPath & strFileName
        
        ' Get the next .doc file
        strFileName = Dir
    Loop
    
    MsgBox "Conversion completed successfully!"
End Sub
