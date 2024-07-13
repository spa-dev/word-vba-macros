# word-vba-macros
A collection of Word VBA macros intended for Medical Writing projects.

## Word Versions
The macros were developed with Office 365 for Windows but should work with earlier versions of Word. The macros have not been tested with Word for Apple Mac.

## Macro Installation
To install macros, open the macro editor in Microsoft Word. In Windows, press Alt+F11 while in a document to open the editor. You can also press Alt+F8 to view current macros. Refer to the [official documentation](https://support.microsoft.com/en-us/office/create-or-run-a-macro-c6b99036-905c-49a6-818a-dfb98b7c3c9c) for additional details.

## List of Macros
Here's a brief description of the content of this repository:

**`AutoFitTablesToMargins`**

A very simple macro to fit all tables in the document to the page margins. It also stops tables from automatically adjusting the column widths to fit the contents, which is often annoying. This is the equivalent to unchecking "Table properties" > "Options" > "Automatically resize to fit contents." If you prefer to keep that behavior, just comment out the relevant line of code.

**`DateCalculator`**

This macro adds a given number of days to a date. Its intended purpose is to help write or QC dates, e.g., in patient safety narratives. It calculates future dates based on the selected text plus the user-provided number of days. The output is displayed in "dd-mmm-yyyy" format. The macro provides both the inclusive and non-inclusive end date.

**`DrugNameReplacer`**

This macro finds a company-specific drug name/code and replaces it with the drug's generic name, applying the correct capitalization based on context. For example, capitalizing the first letter if the name is used at the start of a sentence, as title case or upper case when used in headings, or as lower case within a sentence. Replacement is based on the formatting styles used in the document. It will also exclude replacements within study identifiers (if applicable to your use case). The macro must be modified accordingly per your style guide, drug names, and study identifiers.

The standard Word 'Find and Replace' does a poor job of matching the correct case of a replacement. If you've ever had to find and replace a drug name 300+ times in an Investigator's Brochure, fixing the capitalization as you go, this macro is for you.
  
Note that headers and footers are not checked. You should be using [alignment tabs](https://cybertext.wordpress.com/2014/07/25/word-auto-aligning-headerfooter-info-in-portrait-and-landscape-pages) to ensure you only have to replace it once in the header anyway, using 'Link to Previous', even across portrait and landscape pages.

