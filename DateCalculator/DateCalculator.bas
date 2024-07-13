Attribute VB_Name = "DateCalculator"
Sub DateCalculator()

    ' =========================
    ' Macro created 2024 by spa-dev
    ' GitHub: https://github.com/spa-dev/word-vba-macros
    ' GNU GPL v3 License: https://www.gnu.org/licenses/gpl-3.0.html
    ' =========================
    ' This macro adds a given number of days to a date. It calculates future dates 
    ' based on user-selected text representing a date in a valid format. Note that 
    ' ordinal indicators (e.g., "st", "nd", "rd", etc.) are not considered valid.
    ' The user is prompted to enter the number of days to add to the selected date.
    ' Results are displayed in a message box. The output includes the original date, 
    ' the date after adding the specified days, and the date including one 
    ' additional day (end date inclusive), all formatted as "dd-mmm-yyyy".
    ' Revise the output format as needed (refer to VBA date format documentation).
    ' =========================

    Dim selectedText As String
    Dim startDate As Date
    Dim daysToAdd As Variant ' Used Variant to handle Cancel properly
    Dim result1 As Date
    Dim result2 As Date
    Dim outputMessage As String

    ' Check if text is selected
    If Selection.Type = wdSelectionIP Then
        MsgBox "Please select a date before running this macro.", vbExclamation
        Exit Sub
    End If

    ' Get selected text
    selectedText = Selection.text
    selectedText = Trim(selectedText)

    ' Attempt to convert selected text to a date
    On Error Resume Next
    startDate = DateValue(selectedText)
    On Error GoTo 0

    If startDate = 0 Then
        MsgBox "Selected text is not a valid date or does not contain valid separators.", vbExclamation
        Exit Sub
    End If

    ' Ask user for number of days to add
    daysToAdd = InputBox("Enter the number of days to add to the date:", "Days to Add", 1)

    ' Check if user clicked Cancel
    If daysToAdd = "" Then
        Exit Sub ' Exit gracefully if user clicked Cancel
    End If

    ' Convert daysToAdd to integer
    daysToAdd = Val(daysToAdd)

    ' Calculate the dates
    result1 = DateAdd("d", daysToAdd, startDate)
    result2 = DateAdd("d", daysToAdd + 1, startDate)

    ' Format the output message with "dd-mmm-yyyy"
    outputMessage = "Original Date: " & Format(startDate, "dd-mmm-yyyy") & vbCrLf & _
                    "Added " & daysToAdd & " day(s): " & Format(result1, "dd-mmm-yyyy") & vbCrLf & _
                    "Including end day (+1 day): " & Format(result2, "dd-mmm-yyyy")

    ' Display the results in a message box
    MsgBox outputMessage, vbInformation, "Date Calculation Results"

End Sub

