import xlrd

# I originally wrote this to cross-check two spreadsheets that had contact information in order to find common matches between the two.
# My first spreadsheet had a list of names and contact information for each of the names
# My second spreadsheet were the people I wanted to contact, but I didn't have their contact information.
# Cross-referencing the two spreadsheets on the first and last name fields allowed me to get the contact information for all of the people I wanted to contact.

def CSVtoDict(csv_filename):

    # 2013 Feb 20 Update: this is probably obsolete given the csv module.  The more you know(tm) Still, I'm glad I wrote this as it was a great learning exercise.

    with open(csv_filename) as csv_file:

        header_row = csv_file.readline()
        header_row = header_row.split(",")

        if header_row[len(header_row)-1][-1:] == "\n":
            header_row[len(header_row)-1] = header_row[len(header_row)-1][:-1]
        
        csv = csv_file.read()
        csv = csv.split("\n")

    csv_as_dict = {}.fromkeys(xrange(len(csv)-1))
    
    for row_iter in xrange(len(csv)-1):
        current_row = csv[row_iter].split(",")
        csv_as_dict[row_iter] = {}.fromkeys(header_row)
        for column, col_iter in zip(header_row, xrange(len(header_row))):
            csv_as_dict[row_iter][column] = current_row[col_iter]
          
    return csv_as_dict

def XLStoDict(xls_filename, number_type = "float"):

    " XLStoDict: Reads an XLS or XLSX file and returns a dictionary of the contents. Optional argument: Numbers are floats by default unless you set number_type to 'int'. "

    xls_file = xlrd.open_workbook(xls_filename)
    xls_read = xls_file.sheet_by_index(0)

    xls_as_dict = {}
    header_row = []

    xls_as_dict = {}.fromkeys(xrange(xls_read.nrows-1))

    for col_index in xrange(0, xls_read.ncols):
        if number_type == "int":
            try:
                header_row.append(str(int(xls_read.cell(0, col_index).value)))
            except ValueError:
                header_row.append(xls_read.cell(0, col_index).value)
        else:
            header_row.append(xls_read.cell(0, col_index).value)

    for row_index in xrange(1, xls_read.nrows):
        xls_as_dict[row_index] = {}.fromkeys(header_row)
        for column, col_index in zip(header_row, xrange(0, xls_read.ncols)):
            if number_type == "int":
                try:
                    xls_as_dict[row_index][column] = str(int(xls_read.cell(row_index, col_index).value))
                except ValueError:
                    xls_as_dict[row_index][column] = xls_read.cell(row_index, col_index).value
            else:
                xls_as_dict[row_index][column] = xls_read.cell(row_index, col_index).value

    return xls_as_dict

def CrossReferenceSpreadsheets(spreadsheet_one_as_dict, spreadsheet_two_as_dict, matching_columns):

    " CrossReferenceSpreadsheets: Checks spreadsheet_one against spreadsheet_two on all matching_columns to find value matches.  Yields row number and values found. "

    warned_already = []

    for x in spreadsheet_one_as_dict:
        for y in spreadsheet_two_as_dict:

            all_fields_matched = True

            for column in matching_columns:
                try:
                    if spreadsheet_one_as_dict[x][column] != spreadsheet_two_as_dict[y][column]:
                        all_fields_matched = False
                except KeyError:
                    if column not in warned_already:
                        print "Column '%s' not found in one or both of your spreadsheets.  Ignoring this error and matching on all other like columns." % column
                                        
                    warned_already.append(column)
                    warned_already = list(set(warned_already))  

            if all_fields_matched == True:
                yield (y, (spreadsheet_one_as_dict[x], spreadsheet_two_as_dict[y]))
