# Code 1.2: Extracting Email Metadata of a Folder (Subject, Received Date, Sender) to an Excel File

## Code Description
This VBA script automates the extraction of email metadata from a selected Outlook folder. It retrieves the **Subject**, **Received Date (with time)**, and **Sender Name** of each email and exports this information into an Excel file. The script creates the Excel file in the user’s Desktop and names it after the selected Outlook folder.

## Code Features
- **Folder Selection**: Prompts the user to choose an Outlook folder for metadata extraction.
- **Desktop as Destination**: Automatically saves the Excel file on the user’s Desktop.
- **Metadata Details**: Each email's **Subject**, **Received Date**, and **Sender Name** are included in the Excel file.
- **Naming Consistency**: The Excel file is named after the selected folder, with illegal characters stripped for compatibility.

## Code Prerequisites
- **Microsoft Outlook**: Ensure that Outlook is installed and configured on your system.
- **Microsoft Excel**: Ensure that Excel is installed to create and save the output file.
- **Folder Access**: Verify access to the desired Outlook folder.
- **Sufficient Storage**: Ensure adequate disk space for the Excel file.

## Code Usage Instructions
1. Open Microsoft Outlook.
2. Access the VBA editor and paste the script into a module.
3. Run the script from the VBA editor.
4. Select the Outlook folder you wish to extract metadata from when prompted.
5. Check your Desktop for the Excel file, which will have the same name as the selected Outlook folder.

## Limitations
- **Single Folder Processing**: The script only processes the selected folder. Subfolders are not included in this version.
- **One-Time Execution**: Designed for one-off extractions. For repetitive tasks, consider automating the script further or creating scheduled macros.

## Code Snippet
```vba
Option Explicit

Dim StrSavePath As String

Sub ExtractEmailInfoToExcel_savedtoDesktop()

    Dim i As Long
    Dim j As Long
    Dim StrFolder As String
    Dim iNameSpace As NameSpace
    Dim myOlApp As Outlook.Application
    Dim ChosenFolder As MAPIFolder
    Dim mItem As Object ' Generic Object to handle multiple item types
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim DesktopPath As String
    
    ' Set up Outlook and Excel objects
    On Error GoTo ErrorHandler
    Set myOlApp = Outlook.Application
    Set iNameSpace = myOlApp.GetNamespace("MAPI")
    Set ChosenFolder = iNameSpace.PickFolder
    If ChosenFolder Is Nothing Then GoTo ExitSub

    ' Set Desktop Path for saving the Excel file
    DesktopPath = Environ("USERPROFILE") & "\Desktop\"

    ' Initialize Excel Application
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False ' Keep Excel hidden during the process
    
    ' Create a new workbook and set the sheet
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)
    xlSheet.Name = "Email Log"

    ' Set headers for the Excel file
    xlSheet.Cells(1, 1).Value = "Subject"
    xlSheet.Cells(1, 2).Value = "Received Date"
    xlSheet.Cells(1, 3).Value = "Sender Name"
    
    ' Set folder name as the file name
    StrFolder = StripIllegalChar(ChosenFolder.Name)

    ' Loop through all emails in the selected folder
    For j = 1 To ChosenFolder.Items.Count
        Set mItem = ChosenFolder.Items(j)
        
        ' Only process Mail Items
        If TypeOf mItem Is MailItem Then
            ' Add email details to the Excel sheet
            xlSheet.Cells(j + 1, 1).Value = mItem.Subject
            xlSheet.Cells(j + 1, 2).Value = mItem.ReceivedTime
            xlSheet.Cells(j + 1, 3).Value = mItem.SenderName
        End If
    Next j

    ' Save the Excel workbook with the folder's name on the Desktop
    xlBook.SaveAs DesktopPath & StrFolder & ".xlsx"
    xlBook.Close False
    xlApp.Quit

    GoTo ExitSub

ExitSub:
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume ExitSub
End Sub

' Function to strip illegal characters from filenames
Function StripIllegalChar(StrInput As String) As String
    Dim RegX As Object
    Set RegX = CreateObject("vbscript.regexp")
    
    RegX.Pattern = "[\" & Chr(34) & "\!\@\#\$\%\^\&\*\(\)\=\+\|\[\]\{\}\`\'\;\:\<\>\?\/\,]"
    RegX.IgnoreCase = True
    RegX.Global = True
    
    StripIllegalChar = RegX.Replace(StrInput, "")
End Function
```
