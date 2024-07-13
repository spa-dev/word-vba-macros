Attribute VB_Name = "FormatBoldTextBlue"
Sub FormatBoldTextBlue()

    '=========================
    'Macro created 2023 by spa-dev
    'GitHub: https://github.com/spa-dev/word-vba-macros
    'GNU GPL v3 License: https://www.gnu.org/licenses/gpl-3.0.html
    '=========================

    Dim rng As Range
    Dim char As Variant

    ' Check if text is selected in the document
    If Selection.Type = wdSelectionNormal Then
        ' Set the range to the selected text
        Set rng = Selection.Range
        ' Loop through each character in the selected range
        For Each char In rng.Characters
            ' Check if the character is bold
            If char.Font.Bold = True Then
                ' Set the font color to blue
                char.Font.Color = RGB(0, 0, 255)
                ' To set the color to red, delete above line and uncomment below
                ' char.Font.Color = RGB(255, 0, 0)
            End If
        Next char
    Else
        MsgBox "Please select text before running this macro.", vbExclamation, "No Selection"
    End If
    
End Sub

