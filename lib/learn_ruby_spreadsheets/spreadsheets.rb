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
    cylinderSheet[target_row][5].change_contents("", "SUM(2+2)") Creates a blank cell (this could be replaced with real data), and adds the formula
            I can't seem to get the formula to evaluate for the value of the cell for some reason 
            My solution is to:
            - use Ruby to evaluate the same data that the formula is meant to evaluate
            - pass that data in as the first argument of .change_contents
            - pass the excel formula as the second argument
            This will ensure that the cell value is the same as what the formula would evaluate to while allowing for the formula to
            be copied and pasted directly in the spreadsheet should action be performed.
            

    row.cells returns array-like object of all cells
    row.cells.each lets me itterate over all the cells to pull out values

    worksheet.each iterates over all rows starting from the first
    worksheet.reverse_each iterates over all rows starting from the last
        the rows are still read from left to right

    cylinderSheet.each do |row| This will print out the qty of each row if it has a qty and a line, otherwise the row is skipped
        puts row[qty].value if row && row[qty] && row[line].value  
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

        #Column Identities
        workorder_total_value_a = 1
        workorder_total_cylinders_a = 2
        qty = 3
        customer_or_total = 4
        po_or_averageunit = 6
        sales_employee = 9
        unit_price = 10
        line_price_or_wordorder_total_value_b = 11 #unit price * qty
        model_or_totale_extended = 13
        date = 15
        special_modifiers = 17
        workorder_number = 19
        line = 20
        assy_initials = 21
        quality_initials = 22
        stake_and_seal = 23
        combined_serial_number = 24
        notes = 27
        custom_notes = 28

        # new_empty_row = [cylinderSheet]

        #I need to find the last row which contains a cylinder which has not yet been added to a work order
        #and the last row of the previous work order (which will contain the string "TOTAL #")
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

        #I need to target the row after the last row containing data to add the new total row
        #and black separator line for the new work order
        target_row = last_row_with_data.r 
        cylinderSheet[target_row].cells.each.with_index do |cell, index|
            cell.change_contents(total_row[index].value) if index != qty && index != 10
        end
        blank_row = target_row + 1
        cylinderSheet[blank_row].cells.each.with_index do |cell, index|
            cell.change_contents(cylinderSheet[total_row.r][index].value)
            cell.change_fill("000000") if index > 2 && index < row_size - 1
        end

        #Add the total number of cylinders to the new work order and add the formula to the cell too
        total_workorder_cylinders = 0
        cylinderSheet.each do |row|
            total_workorder_cylinders += row[qty].value if row.r > total_row.r + 1 && row.r <= target_row
        end
        cylinderSheet[target_row][qty].change_contents(total_workorder_cylinders, "SUM(D#{total_row.r + 2}:D#{target_row})")

        #Add a new blank row to the bottom of the spreadsheet
        last_row = cylinderSheet.sheet_data[-1].r
        cylinderSheet.insert_row(last_row)
        cylinderSheet[last_row + 1].cells.each.with_index do |cell, index|
            cell.change_contents(cylinderSheet[last_row][index].value, cell_value.formula)
        end
        
    
        #Save what I've done to a new spreadsheet called test
        #This is to preserve the original for continued testing
        cylinderBook.write("data/test.xlsx")

        # binding.pry
    end
end