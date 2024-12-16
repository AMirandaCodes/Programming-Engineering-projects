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

### Code Snippet
```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/sp/column-formatting.schema.json",
  "elmType": "img",
  "attributes": {
    "src": "@thumbnail.small"
  },
  "style": {
    "display": "block",
    "margin": "0 auto",
    "max-height": "400px"
  },
  "customCardProps": {
    "openOnEvent": "hover",
    "isBeakVisible": true,
    "formatter": {
      "elmType": "img",
      "attributes": {
        "src": "@thumbnail.2000x600"
      },
      "style": {
        "max-width": "10000%",
        "max-height": "1000%"
      }
    }
  }
}
```
