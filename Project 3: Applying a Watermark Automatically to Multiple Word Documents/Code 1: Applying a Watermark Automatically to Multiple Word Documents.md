# Code 1: Applying a Watermark Automatically to Multiple Word Documents

## Code Description
This VBA script automatically applies a specified watermark to the Word documents of a predefined folder.

## Code Features
- **Pre-set watermark and folder selection**: The code need to be updated with the custom fields of the user's watermark and folder, then when the code is ran it automatically applies the watermark as instructed (so the user is not prompted to choose a location)

## Code Usage Instructions
1. Open Word.
2. Open the VBA script from the VBA editor.
3. Replace the following parts of the code with your custom names:
- Replace `"C:\Path\To\Your\Documents\"` with the folder path containing your Word documents.
- Replace `"Watermark Name"` with the exact name of your watermark as saved in the gallery.
- Replace `"C:\Path\To\Watermark Template"` with the full path to your watermark template.
4. Run the VBA script.
5. Check that the documents in that folder have the watermark applied.

## Limitations
- This script is designed to work with the watermark name, watermark template and watermark template file path that is specified in the code. If any of these change, the code needs to be updated.

### Code snippet
```
Sub ApplyCustomWatermarkToDocuments()
    Dim folderPath As String
    Dim fileName As String
    Dim doc As Document
    Dim watermarkName As String
    Dim templatePath As String
    
    ' Customize these variables
    folderPath = "C:\Path\To\Your\Documents\" ' Folder containing the Word documents
    watermarkName = "Watermark Name" ' Name of your custom watermark
    templatePath = "C:\Path\To\Watermark Template" ' Full path to the template containing the watermark
    
    ' Ensure folder path ends with a backslash
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' Get the first file in the folder
    fileName = Dir(folderPath & "*.doc*") ' Looks for .doc and .docx files
    
    ' Load the template containing the custom watermark
    Dim templateDoc As Document
    Set templateDoc = Documents.Open(templatePath, ReadOnly:=True)
    
    ' Process each file in the folder
    While fileName <> ""
        ' Open the document
        Set doc = Documents.Open(folderPath & fileName)
        
        ' Apply the custom watermark
        On Error Resume Next ' Avoid errors if the watermark is already applied
        templateDoc.AttachedTemplate.AutoTextEntries(watermarkName).Insert _
            Where:=doc.Sections(1).Headers(wdHeaderFooterPrimary).Range, _
            RichText:=True
        On Error GoTo 0
        
        ' Save and close the document
        doc.Save
        doc.Close
        fileName = Dir ' Get the next file
    Wend
    
    ' Close the template
    templateDoc.Close SaveChanges:=False
    
    MsgBox "Watermark applied to all documents in the folder.", vbInformation
End Sub
```
