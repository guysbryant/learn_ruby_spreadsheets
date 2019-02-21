=begin
    Working with rubyXL
    workbook = RubyXL::Parser.parse("path/to/xlsx")
    workbook.write("path/to/desired/Excel/file.xlsx")

    work_sheet = workbook[0]
    sheetByName = workbook['sheet_name'] this will find matching sheet

    row1 = work_sheet[0]
    workbook[0][0]

    cell1 = row1[0]
    workbook[0][0][0]
    work_sheet[0][0]

    cell_value = cell1.value to read the cell value
    cell_value.change_contents("new value") to change the cell value the quotes insert data as string, leave off for int
    cell_value.change_contents("new value", cell_value.formula) to preserve the formula

    row.cells returns array-like object of all cells
    row.cells.each lets me itterate over all the cells to pull out values

    worksheet.each iterates over all rows starting from the first
    worksheet.reverse_each iterates over all rows starting from the last
        the rows are still read from left to right

    cyinderSheet.each do |row| This will print out the qty of each row if it has a qty and a serial, otherwise the row is skipped
        puts row[qty].value if row && row[qty] && row[serial].value  
    end
=end
    
class LearnRubySpreadsheets::SpreadSheets

    def initialize
        workbook = RubyXL::Parser.parse("data/first_test_spreadsheet_data.xlsx")
        worksheet = workbook[0]
        cylinderBook = RubyXL::Parser.parse("data/Scrubbed Data.xlsx")
        cylinderSheet = cylinderBook[0]

        qty = 3
        serial = 19
        line = 20
        assy_initials = 21
        quality_initials = 22
        stake_and_seal = 23

        customer_or_total = 4
        po_or_averageunit = 6
        
        binding.pry
    end
end