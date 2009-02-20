########################################################################
# example_basic.rb
# 
# Simply run this and then open up basic_example.xls with MS Excel
# (or Gnumeric or whatever) to view the results.
########################################################################
base = File.basename(Dir.pwd)
if base == "examples" || base =~ /spreadsheet/i
   Dir.chdir("..") if base == "examples"
   $LOAD_PATH.unshift(Dir.pwd)
   $LOAD_PATH.unshift(Dir.pwd + "/lib")
   Dir.chdir("examples") rescue nil
end

require "spreadsheet/excel"
include Spreadsheet

puts "VERSION: " + Excel::VERSION

workbook = Excel.new("basic_example.xls")

worksheet = workbook.add_worksheet

worksheet.write(0, 0, "Hello")
worksheet.write(1, 0, ["Matz","Larry","Guido"])

workbook.close
