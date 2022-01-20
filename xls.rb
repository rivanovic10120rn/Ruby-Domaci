require 'roo-xls'

class XLS
    include Enumerable
    attr_accessor :excel, :headers, :full, :row

    def initialize(root)
        @excel = Roo::Spreadsheet.open(root)
        @excel = Roo::Excelx.new(root,{:expand_merged_ranges => false})
        initTable
    end

    def initTable()
        
        @headers = Hash[]
        @full = []
        @row = Array.new

        @excel.sheets.each do |currentSheet|

            if (excel.sheet(currentSheet) != nil)

                sheet = excel.sheet(currentSheet)
                
                if(sheet.first_row != nil)

                    first_row = sheet.first_row
                    first_column = sheet.first_column
                    puts "First row and column set\n"

                    last_row = sheet.last_row
                    last_column = sheet.last_column
                    puts "Last row and column set\n"

                    
                    currentHeader = ""
                    first_column.upto(last_column) do |pointrc|

                        fraction = Array.new

                        first_row.upto(last_row) do |pointrr|

                            currentCell = sheet.cell(pointrr,pointrc)
                            if (currentCell != nil)
                                if(pointrr == first_row)
                                    currentHeader = currentCell
                                    headers[currentHeader] = nil
                                else
                                    fraction << (currentCell)
                                end
                            end

                        end
                        headers[currentHeader] = fraction
                    end
                    puts "Table Headers Matrix ready\n"

                    counter = 0
                    first_row.upto(last_row) do |pointrr|

                        twodimensional = []

                        first_column.upto(last_column) do |pointrc|

                            currentCell = sheet.cell(pointrr,pointrc)
                            if (currentCell != nil)
                                twodimensional << (currentCell)
                            end

                        end
                        full[counter] = *twodimensional
                        counter+=1
                    end
                    puts "Table Rows Matrix ready\n"

                else
                    puts "Sheet is empty\n"
                end
            end
        end
    end

    def row(index)
        @row = full[index]
    end

    def each(&block)
        @full.each(&block)
    end

    def header_value(c, m, &b)
        c.class_eval {
            define_method(m, &b)
        }
    end
    
    def header_search_methods
        self.headers.each do |key, value|
            header_value(XLS, key) do
                value
            end
        end    
    end

end

class Fraction < Array
    # puts "Table Rows Matrix ready\n"
                    # first_row.upto(last_row) do |pointrr|
                    #     first_column.upto(last_column) do |pointrc|
                    #         currentCell = sheet.cell(pointrr,pointrc)
                    #         if (currentCell != nil)
                    #             if(currentCell == first_row)
                    #                 rows.push(currentCell)
                    #             else
                    #                 fraction.push(currentCell)
                    #             end
                    #         end
                    #     end
                    # end
end