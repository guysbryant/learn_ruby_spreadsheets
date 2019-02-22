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

    .r Gives the actual row which are numbered starting from 1
    cylinderSheet[index] is how I manipulate rows and it is numbered starting from 0
    target_row = last_row_with_data.r This will let me target a specific row
        cylinderSheet[target_row].cells.each.with_index do |cell, index| This will let me copy the values from each cell one row into the cells of another row
            cell.change_contents(last_row_with_data[index].value)
        end
=end
    
class LearnRubySpreadsheets::SpreadSheets

    def initialize
        workbook = RubyXL::Parser.parse("data/first_test_spreadsheet_data.xlsx")
        worksheet = workbook[0]
        cylinderBook = RubyXL::Parser.parse("data/Scrubbed Data.xlsx")
        cylinderSheet = cylinderBook[0]
        row_size = cylinderSheet[1].size

        qty = 3
        serial = 19
        line = 20
        assy_initials = 21
        quality_initials = 22
        stake_and_seal = 23

        customer_or_total = 4
        po_or_averageunit = 6

        total_row = nil
        last_row_with_data = nil

        cylinderSheet.reverse_each do |row|
            if row && row[qty] && row[qty].value
                last_row_with_data = row if last_row_with_data == nil
            end
            if row[customer_or_total].value == "TOTAL #"
                total_row = row if total_row == nil
            end
            break if total_row != nil && last_row_with_data != nil
        end

        target_row = last_row_with_data.r 
        cylinderSheet[target_row].cells.each.with_index do |cell, index|
            cell.change_contents(total_row[index].value)
        end
        blank_row = target_row + 1
        cylinderSheet[blank_row].cells.each.with_index do |cell, index|
            cell.change_contents(cylinderSheet[total_row.r][index].value)
            cell.change_fill("000000") if index > 2 && index < row_size - 1
        end

    cylinderBook.write("data/test.xlsx")

        binding.pry
    end
end