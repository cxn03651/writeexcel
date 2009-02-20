########################################################################
# example_formula.rb
# 
# Simply run this and then open up format_example.xls with MS Excel
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

version = Excel::VERSION
puts "VERSION: " + version

workbook = Excel.new("format_example.xls")

# Preferred way to add a format
f1 = workbook.add_format(:color=>"blue", :bold=>1, :italic=>true)

# Another way to add a format
f2 = Format.new(
   :color  => "red",
   :bold   => 1,
   :italic => true
)
workbook.add_format(f2)

# Yet another way to add a format
# A tiny bit more overhead doing it this way
f3 = Format.new{ |f|
   f.color  = "green"
   f.bold   = 1
   f.italic = true
}
workbook.add_format(f3)

f4 = Format.new(:num_format => "d mmm yyyy")
f5 = Format.new(:num_format => 0x0f)
workbook.add_format(f4)
workbook.add_format(f5)

worksheet1 = workbook.add_worksheet
worksheet2 = workbook.add_worksheet("number")
worksheet3 = workbook.add_worksheet("text")

worksheet1.write(0, 0, version)
worksheet1.write(0, 1, "Hello", f1)
worksheet1.write(1, 1, ["Matz","Larry","Guido"])

worksheet2.write_column(1, 1, [[1,2,3],[4,5,6],[7,8,9]])
worksheet2.write(0, 0, 8888, f2)

worksheet3.write(0, 0, 36892.521, f5)

worksheet1.format_row(4..5, 30, f1)

worksheet2.format_column(3..4, 25, f2)

worksheet1.write(5,0,"This should be blue")

worksheet2.write(0,4,"This should be red")

workbook.close
