xref-spreadsheets
=================

Python script to cross-reference spreadsheets on like columns to find matching values in all columns given.

I originally wrote this to cross-check two spreadsheets that had contact information in order to find common matches between the two.
My first spreadsheet had a list of names and contact information for each of the names
My second spreadsheet were the people I wanted to contact, but I didn't have their contact information.
Cross-referencing the two spreadsheets on the first and last name fields allowed me to get the contact information for all of the people I wanted to contact.

Examples:

spreadsheet_one_as_dict = CSVtoDict(file1)
spreadsheet_two_as_dict = CSVtoDict(file2)

or

spreadsheet_one_as_dict = XLStoDict(file3, "int")
spreadsheet_two_as_dict = XLStoDict(file4, "int")

then:

print list(CrossReferenceSpreadsheets(spreadsheet_one_as_dict, spreadsheet_two_as_dict, ["pawn", "rook", "queen", "sddd"]))

Quick notes:
* Use "int" as an optional second parameter in XLStoDict when you don't want numbers to be assumed to be floats.
* The column "sddd" doesn't exist in (one or both) spreadsheets, so the program warns you (once!) and moves on, matching on all other columns.
* Values must match in all columns you provide in both spreadsheets or CrossReferenceSpreadsheets will not yield a match.
