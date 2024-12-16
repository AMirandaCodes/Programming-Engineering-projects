# Code 1.3: Extracting Email Metadata of Subfolders (Subject, Received Date, Sender) to Excel Files on Desktop

## Code Description
This VBA script automates the extraction of email metadata (Subject, Received Date with time, and Sender Name) from all subfolders of a selected Outlook folder. For each subfolder, an individual Excel file is created and saved in a designated "Email Info" folder on the user’s Desktop.

## Code Features
- **Folder Selection**: Enables the user to choose an Outlook folder for subfolder metadata extraction.
- **Desktop Destination**: Automatically creates a "Email Info" folder on the Desktop to store the Excel files.
- **Recursive Extraction**: Processes all subfolders within the selected folder.
- **Detailed Metadata**: Captures each email’s **Subject**, **Received Date** (including time), and **Sender Name**.
- **Completion Notification**: Displays a confirmation message once the extraction is complete.

![image](https://github.com/user-attachments/assets/899bd407-7b5f-432a-a93f-a78c1b590b5a)

## Code Prerequisites
- **Microsoft Outlook**: Ensure Outlook is installed and configured on your system.
- **Folder Access**: Verify access to the desired folders in Outlook.
- **Sufficient Storage**: Ensure adequate disk space for the exported Excel files.

## Code Usage Instructions
1. Open Microsoft Outlook.
2. Access the VBA editor and paste the script into a module.
3. Run the script from the VBA editor.
4. Select the Outlook folder whose subfolders you wish to extract email metadata from.
5. Check your Desktop for the "Email Info" folder containing the Excel files.

## Limitations
- **No Metadata for Selected Folder**: Only subfolders are processed; emails in the selected folder itself are not included.
- **One-Off Execution**: Designed for single-use exports.

## Code Snippet
```vba
Option Explicit

Dim StrSavePath As String

Sub ExtractEmailInfoFromSubfolders()
    Dim iNameSpace As NameSpace
    Dim myOlApp As Outlook.Application
    Dim ChosenFolder As MAPIFolder
    Dim SubFolder As MAPIFolder
    Dim DesktopPath As String
    
    ' Set up Outlook objects
    On Error GoTo ErrorHandler
    Set myOlApp = Outlook.Application
    Set iNameSpace = myOlApp.GetNamespace("MAPI")
    Set ChosenFolder = iNameSpace.PickFolder
    If ChosenFolder Is Nothing Then GoTo ExitSub

    ' Set Desktop path and create "Email Info" folder
    DesktopPath = Environ("USERPROFILE") & "\Desktop\Email Info\"
    If Dir(DesktopPath, vbDirectory) = "" Then
        MkDir DesktopPath
    End If

    ' Loop through each subfolder and export email details
    For Each SubFolder In ChosenFolder.Folders
        ExportEmailsToExcel SubFolder, DesktopPath
    Next SubFolder

    MsgBox "Export completed successfully.", vbInformation

    GoTo ExitSub

ExitSub:
    Set myOlApp = Nothing
    Set iNameSpace = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume ExitSub
End Sub

Sub ExportEmailsToExcel(ByVal Folder As MAPIFolder, ByVal SavePath As String)
    Dim mItem As Object
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim j As Long
    Dim FolderName As String

    ' Set up Excel objects
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)
    xlSheet.Name = "Email Log"

    ' Set headers for the Excel sheet
    xlSheet.Cells(1, 1).Value = "Subject"
    xlSheet.Cells(1, 2).Value = "Received Date"
    xlSheet.Cells(1, 3).Value = "Sender Name"

    ' Clean up folder name for file use
    FolderName = StripIllegalChar(Folder.Name)

    ' Loop through emails in the subfolder
    j = 1
    For Each mItem In Folder.Items
        If TypeOf mItem Is MailItem Then
            xlSheet.Cells(j + 1, 1).Value = mItem.Subject
            xlSheet.Cells(j + 1, 2).Value = mItem.ReceivedTime
            xlSheet.Cells(j + 1, 3).Value = mItem.SenderName
            j = j + 1
        End If
    Next mItem

    ' Save Excel workbook with subfolder's name
    xlBook.SaveAs SavePath & FolderName & ".xlsx"
    xlBook.Close False
    xlApp.Quit

    ' Release Excel objects
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
End Sub

Function StripIllegalChar(StrInput As String) As String
    Dim RegX As Object
    Set RegX = CreateObject("vbscript.regexp")
    
    RegX.Pattern = "[\" & Chr(34) & "\!\@\#\$\%\^\&\*\(\)\=\+\|\[\]\{\}\`\'\;\:\<\>\?\/\,]"
    RegX.IgnoreCase = True
    RegX.Global = True
    
    StripIllegalChar = RegX.Replace(StrInput, "")
End Function
