#  🏭� Plant Registry – Equipment Checker (Excel VBA Automation)

This project is an Excel-based automation tool that scans equipment records and highlights items that are overdue for maintenance.  
It uses **VBA macros** to compare dates, identify equipment older than a defined threshold, and visually highlight those rows for quick action.

---

## 📌 Project Overview
Many organizations maintain logs of equipment, plant assets, and tools in Excel. Manually checking which items are overdue for maintenance is time-consuming and error-prone.

This tool solves that by:

- Automatically scanning each equipment row  
- Comparing **Last Checked Date** with today  
- Highlighting overdue items if **>120 days**  
- Giving users a quick visual summary  

This improves decision-making, auditing, and ensures equipment safety.

---

## 🚀 Features

✔️ Automated equipment health check  
✔️ Macro-enabled Excel workbook (`.xltm`)  
✔️ Color-coded highlighting for overdue items  
✔️ Simple one-button execution  
✔️ Clear, readable VBA code  
✔️ Works on both Windows & Mac Excel  

---

## 📂 Files in This Repository

| File | Description |
|------|-------------|
| **Plant_Registry_Data.xltm** | Main Excel workbook containing the VBA macro |
| **VBA Output.pdf** | Screenshot of the tool and code execution UI |
| **README.md** | Documentation for this project |

---

## 🛠 VBA Logic Summary

The macro loops through each row, checks the date in the "Last Checked" column, and highlights overdue records:

```vb
Sub CheckEquipment()

    'Checks the equipment by scanning cells from A2:GX, one row at a time
    'If the date in column G is more than 4 months old, highlights the row

    Dim currentCell As Range
    Set currentCell = Range("A2")

    Dim currDate As Date
    currDate = Date

    Dim CountEquipment As Integer
    CountEquipment = 0

    'Reset formatting
    Range(currentCell, currentCell.End(xlDown).Offset(0, 6)).Select
    With Selection
        .Interior.Pattern = xlNone
        .Font.Bold = False
        .Font.Underline = xlUnderlineStyleNone
    End With

    'Loop through rows
    While currentCell <> ""

        If currDate - currentCell.Offset(0, 6).Value > 120 Then
        
            CountEquipment = CountEquipment + 1
            
            With currentCell.EntireRow
                .Interior.ColorIndex = 36  'Highlight color
                .Font.Bold = True
            End With
        
        End If

        Set currentCell = currentCell.Offset(1, 0)
    Wend

    MsgBox "Highlighted " & CountEquipment & " items.", vbInformation, "Done!"

End Sub
