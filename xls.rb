require 'roo-xls'

class XLS
    # include Enumerable
    attr_accessor :excel, :headers, :full, :row, :rows

    def initialize(root)
        @excel = Roo::Spreadsheet.open(root)
        initTable
    end

    def initTable()
        
        @headers = Hash[]
        @rows = Hash[]
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

                    #bolje da sam radio sa to_matrix koji bi mi napravio celu matricu kao objekat al sad kasno...
                    
                    currentHeader = ""
                    currentRowIndex =""
                    first_column.upto(last_column) do |pointrc|

                        fraction = Fraction.new

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

                    # index = 0
                    # delete = 0
                    # first_row.upto(last_row) do |pointrc|

                    #     fraction = Fraction.new

                    #     first_column.upto(last_column) do |pointrr|

                    #         currentCell = sheet.cell(pointrr,pointrc)
                    #         if (currentCell != nil)
                    #             if(pointrr == first_column)
                    #                 rows[index] = nil
                    #             else
                    #                 if(currentCell.casecmp "Subtotal" || currentCell.casecmp "Total")
                    #                     delete = index
                    #                 else
                    #                     fraction << (currentCell)
                    #                 end
                    #             end
                    #         end

                    #     end
                    #     rows[index] = fraction
                    #     index +=1
                    # end
                    # puts "Table Rows Matrix ready\n"

                    counter = 0
                    first_row.upto(last_row) do |pointrr|

                        twodimensional = []

                        first_column.upto(last_column) do |pointrc|

                            currentCell = sheet.cell(pointrr,pointrc)
                            if (currentCell != nil)
                                twodimensional << (currentCell)
                            end

                        end
                        # if(!twodimensional.empty?)
                            full[counter] = *twodimensional
                        # end
                        counter+=1
                    end
                    puts "Table Full Matrix ready\n"

                else
                    puts "Sheet is empty\n"
                end
            end
        end
    end

    def delete(index)
        row(index).each do
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

    def row_value(c, m, &b)
        c.class_eval {
            define_method(m, &b)
        }
    end
    
    def rows_search_methods
        self.full.each do |key, value|
            row_value(XLS, key) do
                if key != nil
                    value
                end
            end
        end    
    end

    def +(outsider)
        if self.full[0] == outsider.full[0]
            return self.full+outsider.full[1...]
        end
    end
    def -(outsider)
        if self.full[0] == outsider.full[0]
            return self.full-outsider.full[1...]
        end
    end


end

class Fraction < Array
    def sum
        total=0
        self.each do |cell|
            if(cell != nil and cell.is_a? (Integer))
                total = total + cell.to_i
            end
        end
        total
    end
end