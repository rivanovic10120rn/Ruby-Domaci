require 'roo'
require_relative 'xlsx'
# require_relative 'xls'

# xlsx = Roo::Spreadsheet.open('./excel-tests/RubyTest.xlsx')
# xlsx = Roo::Excelx.new("./excel-tests/RubyTest.xlsx")

# puts xlsx.info
run = XLSX.new('./excel-tests/RubyTest.xlsx')
run2 = XLSX.new('./excel-tests/RubyTest2.xlsx')

puts "Ispis po headerima"
p run.headers
puts
puts "Ispis cele kao dvodimenzionalni niz"
p run.full
puts
puts "Ispis odabranog reda"
p run.row(2)
puts
p run2.headers
puts
puts "Ispis cele kao dvodimenzionalni niz sa implementiranim obracanjem paznje na mergeovane celije"
p run2.full
puts
puts "Ispis odabranog reda i odredjenog elementa unutar tog reda"
p run2.row(4)[0]
puts
puts "Implementacija enumerable"
run.each do |cell|
    p cell
end
puts
puts "Ispis odabranog headera(kolone) kao i odredjenog elementa unutar te kolone"
p run.headers["Second"]
p run.headers["Second"][3]
puts
puts "Ispis odabranog headera u vidu metode kao i odredjenog elementa unutar te kolone"
run.header_search_methods
p run.Second
p run.Second[0]