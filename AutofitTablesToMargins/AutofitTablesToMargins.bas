Attribute VB_Name = "AutoFitTablesToMargins"

Sub AutoFitTablesToMargins()

    Dim tbl As Table

    For Each tbl In ActiveDocument.Tables
        ' Autofit table to page margins
        tbl.AutoFitBehavior (wdAutoFitWindow)
        ' Prevent automatic adjustment of column width to fit contents
        tbl.AllowAutoFit = False
    Next tbl

End Sub
