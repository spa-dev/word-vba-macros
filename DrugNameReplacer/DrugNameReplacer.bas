Attribute VB_Name = "DrugNameReplacer"

Option Explicit

Sub DrugNameReplacer()

    ' =========================
    ' Macro created 2023 by spa-dev
    ' GitHub: https://github.com/spa-dev/word-vba-macros
    ' GNU GPL v3 License: https://www.gnu.org/licenses/gpl-3.0.html
    ' =========================
    ' This macro finds a company-specific drug name/code and replaces it with the generic name,
    ' applying context-sensitive capitalization. It excludes replacements within study identifiers.
    ' Replacement is based on the text styles used in the document. Ensure the styles chosen below 
    ' are consistent with those in your company style guide. Headers and footers are not checked.
    ' =========================
    
    Dim drugCode As String
    Dim studyCode As String
    Dim lowercaseName As String
    Dim titlecaseName As String
    Dim uppercaseName As String
    Dim foundTextStyle As String
    Dim Msg As String

    ' Replace the following names as applicable to your drug and study identifier.
    ' If your studies are not in this format (or similar), comment out the relevant code.
    ' I have no relationship to this company or drug. It's just a good example.

    drugCode = "IMMU-132"
    studyCode = "IMMU-132-01"
    lowercaseName = "sacituzumab govitecan"
    
    ' Use VBA functions to convert to title case and upper case
    titlecaseName = StrConv(lowercaseName, vbProperCase)
    uppercaseName = UCase(lowercaseName)

    ' Show pop-up message. Stop if user does not click Yes.
    Msg = "This macro finds a pre-specified company drug code and replaces it " & _
        "with the generic name with context-sensitive capitalization. " & _
        "It will ignore the drug code when used in a study identifier (if applicable). " & vbCr & vbCr & _
        "This process may take several minutes. " & _
        "Tracked changes will be turned off. " & _
        "Instead, save your original document first then compare the result." & vbCr & vbCr & _
        "Do you want to continue?"

    If MsgBox(Msg, vbYesNo + vbQuestion, "Confirm Replacement") <> vbYes Then
        Exit Sub
    End If

    ActiveDocument.TrackRevisions = False
    Application.ScreenUpdating = False

    ' Replaces the study identifier with uncommon characters to hide it from the routines below.
    ' If you need to hide anything else, add similar code here.
    ' For example, you could hide "i.e." from the sentence detector, wdTitleSentence.
    '
    With ActiveDocument.Content.Find
            .ClearFormatting: .Replacement.ClearFormatting
            .Execute findText:=studyCode, ReplaceWith:="@@!@@!", Replace:=wdReplaceAll
    End With

    ' Set Selection to start of document; also enables monitoring of progress via the scroll bar

    ActiveDocument.Range(0, 0).Select

    With Selection.Find
    .ClearFormatting: .Wrap = wdFindContinue: .Forward = True: .Format = False: .MatchCase = False: .MatchWildcards = False: .text = drugCode: .Execute

        While .Found
            ' Find style of current selection
            foundTextStyle = Selection.Style.NameLocal
            ' Apply correct case for the specified style
            Select Case foundTextStyle
                ' Adjust the styles below per your style guide. Use the exact style names from Word.
                Case "TOC Title", "Title", "Heading 1"
                    Selection.text = uppercaseName
                Case "Table Cell Heading 10pt", "Table Cell Heading 12pt", "Caption", "Heading 2", "Heading 3", "Heading 4", "Heading 5", _
                "Heading 6", "Heading 7", "Heading 8", "Appendix", "Appendix Heading 1", "Appendix Heading 2", "Appendix Heading 3"
                    Selection.text = titlecaseName
                Case Else
                    ' For all other styles, use lowercase
                    Selection.text = lowercaseName
                    ' Change text to sentence case if at the start of a sentence
                    Selection.Range.Case = wdTitleSentence
            End Select

            Selection.Collapse Direction:=wdCollapseEnd
            .Execute
        Wend
    End With

    'Replace the uncommon characters with the study code, to reveal the study code we hid earlier'

    With ActiveDocument.Content.Find
        .ClearFormatting: .Replacement.ClearFormatting
        .Execute findText:="@@!@@!", ReplaceWith:=studyCode, Replace:=wdReplaceAll
    End With

    Application.ScreenUpdating = True

    ' Provide pop-up box summary.
    MsgBox "Replacement complete. " & _
    "Carefully check the results." & vbCrLf & vbNewLine & _
    "Note: Capitalization may be incorrect if not formatted per company styles." & vbNewLine & _
    "Note: Document headers and footers are not checked." & vbNewLine & _
    "Note: The macro identifies the period in e.g. or i.e. as the end of a sentence if no comma is present. Use commas, e.g., like this sentence.", vbInformation, "Replacement Summary"

End Sub


