Excel-to-DSC
============
Use the attached Excel file to create an EAD container list. The process supports creating controlled vocabularies, mixed-content elements (such as title, persname, emph elements, etc.), date elements with ISO-8601 date normalizations, and a hierarchy up to 12 levels deep. Any of the column values may also be repeated now -- resulting in multiple unitdate values at any level of description, for instance -- by adding a new row and setting its level equal to 0.

To get started:

* enter a container list in the attached Excel file, following the examples provided (hiding any of the 55 columns that you won't need to utilize)
* save the Excel file as an "XML Spreadsheet 2003" document
* transform your XML with the "Excel-to-DSC" XSLT style sheet

Or, to convert an EAD file to Excel:
* transform your EAD file(s) with the "DSC-to-Excel" XSLT stylesheet
* Open the ***-excel.xml file(s) with Excel
* Note that there's a third worksheet in your spreadsheet.  If you'd like the collection-level infomration to be merged with your newly-edited container list upon converting the Excel file back into EAD, just make sure that the original EAD filename is available in the same directory and has a filename that's equal to its "eadid" value.
