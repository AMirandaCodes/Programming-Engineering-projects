# Project 1: Exporting Folders from Outlook

## Project Brief
The company requires an automated solution to export all Outlook Public Folders to a local drive. Public Folders are shared email storage locations accessible to all users within the organization. This one-off extraction ensures proper data backup and organization.

## Project Requirements
- **Folder structure**: The original structure of the Public Folders must remain intact in the exported location.
- **Email format**: All emails must be exported in `.msg` format to preserve metadata and compatibility.
- **Scope**: All emails across all Public Folders must be included in the extraction.
- **Execution**: This is a one-off extraction.

## Project Deliverables
- A VBA code solution capable of:
  1. Extracting all Public Folders to a desired location on the local drive.
  2. Preserving the original folder structure.
  3. Exporting emails in `.msg` format.
 
## Features
1. Exporting all folders to a specified drive location.
2. Exporting folders directly to the desktop.
3. Extracting email metadata (e.g., Subject, Sent Date, Sender) to Excel or Notepad.

## Instructions
1. Open Outlook and ensure access to all necessary Public Folders.
2. Run the VBA script following the provided instructions.
3. Verify the exported files in the target location.

### Subprojects and Codes
- **Code 1.1.1**: Exporting all folders to a local drive.
- **Code 1.1.2**: Exporting folders to the Desktop.
- **Code 1.2.1**: Exporting email metadata to Excel.
- **Code 1.2.2**: Exporting email metadata to Notepad.
