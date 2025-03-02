# excel_automation
I encountered a challenge while working with a large dataset in Excel that had extra spaces in multiple columns, including headers. These unwanted spaces were causing inconsistencies in reporting and analysis, reporting, and automation workflows. Manually cleaning such data is inefficient and time-consuming.


**Problem sTATEMENT**:
Manually cleaning thousands of rows wasn’t an efficient solution. Using Excel’s TRIM function on a large dataset was time-consuming, and I needed a more scalable approach.
1. Incorrect filtering, sorting, and lookups (e.g., VLOOKUP, INDEX-MATCH).
2. Mismatched data in reports and dashboards.
3. Performance issues when working with large datasets.

**Solution:**
I automated the process using VBA (Visual Basic for Applications)! A simple VBA script helped me remove extra spaces across the entire dataset in seconds. Now, my data is cleaner, more accurate, and ready for analysis!

This VBA script:
✔ Loops through all columns and rows in the dataset.
✔ Removes leading, trailing, and extra spaces between words.
✔ Cleans data instantly, making it ready for accurate analysis.

🎥 I’ve documented my approach in a short video along with the dataset to demonstrate the solution in action.

**Key Takeaways:**
Automating repetitive tasks saves time and improves efficiency.
VBA is a powerful tool for data transformation in Excel.
Clean data ensures accurate insights and better decision-making.


**VBA Script**
Sub RemoveExtraSpacesWithHighlight()
    Dim ws As Worksheet
    Dim cell As Range
    Dim originalValue As String
    Dim cleanedValue As String
    Dim changeCount As Long
    
    changeCount = 0  ' Initialize change counter

    ' Loop through all sheets (optional, remove if needed)
    For Each ws In ActiveWorkbook.Sheets
        ' Loop through all used cells
        For Each cell In ws.UsedRange
            If Not IsEmpty(cell.Value) Then
                originalValue = cell.Value
                cleanedValue = Application.WorksheetFunction.Trim(originalValue)
                
                ' Check if cleaning made a difference
                If originalValue <> cleanedValue Then
                    cell.Value = cleanedValue  ' Update cell with cleaned text
                    cell.Interior.Color = RGB(255, 0, 0)  ' Highlight in red
                    changeCount = changeCount + 1
                End If
            End If
        Next cell
    Next ws

    ' Display number of changes made
    MsgBox "Extra spaces removed and highlighted in red for " & changeCount & " cells!", vbInformation, "Cleanup Complete"
End Sub
