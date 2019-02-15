=begin
    Working with rubyXL
    workbook = RubyXL::Parser.parse("path/to/xlsx")

    work_sheet = workbook[0]
    sheetByName = workbook['sheet_name'] this will find matching sheet

    row1 = work_sheet[0]
    workbook[0][0]

    cell1 = row1[0]
    workbook[0][0][0]
    work_sheet[0][0]

    cell_value = cell1.value
=end
    
end
class LearnRubySpreadsheets::SpreadSheets

    def initialize
        workbook = RubyXL::Parser.parse("data/first_test_spreadsheet_data.xlsx")
        worksheet = workbook[0]
        binding.pry
    end
end