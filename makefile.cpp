Option Explicit

Sub UpdateWiOnButtonClick()
    ' Declare variables
    Dim newPart As Worksheet
    Dim wiNew As Worksheet
    Dim newPartRow As Long
    Dim newName As String
    Dim newWorkbook As Workbook
    Dim targetSheet As Worksheet

    ' Set the worksheet references
    Set newPart = ThisWorkbook.Sheets("Part Information") ' Change to your actual sheet name
    Set wiNew = ThisWorkbook.Sheets("Work Instruction") ' Change to your actual sheet name

    ' Ask the user to input the part row they want to copy
    On Error Resume Next
    newPartRow = Application.InputBox("Enter the part row number", Type:=1)
    On Error GoTo 0

    ' Check if the user canceled the input
    If newPartRow = 0 Then
        MsgBox "Operation canceled.", vbExclamation
        Exit Sub
    End If

    ' Check if the entered row number is valid
    If newPartRow < 1 Or newPartRow > newPart.Cells(newPart.Rows.Count, "B").End(xlUp).Row Then
        MsgBox "Invalid row number.", vbExclamation
        Exit Sub
    End If

    ' Copy data from "Part Information" sheet to "Work Instruction" sheet
    wiNew.Range("C3").Value = newPart.Range("B" & newPartRow).Value
    wiNew.Range("C4").Value = newPart.Range("C" & newPartRow).Value
    ' ... (update the rest of the code accordingly)

    ' Generate a new name based on the values in the chosen row of "Part Information" sheet
    newName = newPart.Cells(newPartRow, "B").Value & "-" & newPart.Cells(newPartRow, "C").Value

    ' Create a new workbook
    Set newWorkbook = Workbooks.Add

    ' Copy the "Work Instruction" sheet from the existing workbook to the new workbook
    wiNew.Copy Before:=newWorkbook.Sheets(1)

    ' Rename the copied sheet in the new workbook
    Set targetSheet = newWorkbook.Sheets(1)
    targetSheet.Name = "Work Instruction"

    ' Save the new workbook with the generated name
    newWorkbook.SaveAs "C:\Users\User\Desktop\WI\" & newName & ".xlsx"

    ' Close the new workbook (optional, depending on your needs)
    newWorkbook.Close SaveChanges:=False

    MsgBox "New workbook created and saved successfully!", vbInformation
End Sub
