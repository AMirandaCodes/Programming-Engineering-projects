# Project 1: Customising a Hover Card in SharePoint

## Project Brief
The company uses SharePoint Online and requires an efficient method for quickly previewing documents without opening them. While the primary focus is on previewing `.msg` email files in a specific library, the solution should be adaptable to other file types and applicable across any SharePoint library.

## Project Requirements
- **Clear Document Preview**: The preview must display a clear and legible version of the document for quick review.
- **Minimal User Actions**: The preview should be triggered with minimal effort, such as hovering over the file name or performing a single click.
- **Library-Wide Availability**: The preview functionality must be accessible for all folders and documents within the specified library.

## Project Deliverables
- A working solution that enables:
  1. A user-friendly and visually clear preview for documents.
  2. Consistent functionality across all folders and file types within the targeted SharePoint library.
- The solution may utilise SharePoint's default features (e.g., hover cards, file viewers) or a custom implementation if necessary.
  
## Additional Considerations
- **Compatibility**: Ensure the solution works seamlessly across common web browsers and devices.
- **Performance**: Avoid introducing significant delays or performance issues when accessing the document library.
- **Scalability**: While designed for one library, the solution should be easily replicable for other libraries if needed.
- **Focus on `.msg` Files**: Ensure `.msg` files are prioritised in testing, as they are the primary file type for this project.

# Code 1: Hover Card in SharePoint (Email Focused)

## Code Description
This solution combines a SharePoint feature with JSON formatting to customise the behaviour of the **Thumbnail** column. It leverages the built-in Thumbnail column available in SharePoint, which provides a small preview of files. The JSON formatting is applied to enlarge the thumbnail dynamically when a user hovers over it, making email `.msg` files more legible.

## Code Features
- **Customisable Preview Size**: The hover card displays an enlarged version of the thumbnail, with dimensions that can be adjusted as needed.
- **Hover Trigger**: The hover card is activated only when the mouse pointer is placed over the thumbnail in the Thumbnail column, preventing accidental previews.

![image](https://github.com/user-attachments/assets/a23cab66-2d89-4f6e-b328-315d2a427649)

## Code Usage Instructions
1. **Create the Thumbnail Column**:
   - Go to the intended document library in SharePoint.
   - Create a new column with the exact name **Thumbnail**.
   - Choose **Single Line of Text** or **Picture** as the column type.

2. **Display Thumbnails**:
   - SharePoint will automatically generate a small thumbnail preview for each file in this column.

3. **Apply JSON Formatting**:
   - Click on the **Thumbnail** column header.
   - Navigate to **Column Settings > Format this column > Format columns**.
   - Copy and paste the provided JSON code snippet into the editor.

4. **Save and Test**:
   - Click **Save** to apply the changes.
   - Hover over the thumbnails in the column to test the hover card functionality.

## Limitations
- **Email Length**: Some `.msg` files with lengthy email chains may exceed the hover card's default dimensions, making them harder to read. Adjust the card size if needed.
- **File Type Adjustments**: Non-email file types may require tweaking the hover card's dimensions for optimal display.
- **Images in emails not visible**: Any images embedded in the email are not visible from the hover card (i.e., they are not loaded, and they are represented with an 'X' instead)
