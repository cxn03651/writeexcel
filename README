== Description
   This is a port of John McNamara's Perl module "Spreadsheet::WriteExcel". It
   allows you to generate Microsoft Excel compatible spreadsheets (in 
   Excel 95 format) on *any* platform. These spreadsheets are viewable
   with most other popular spreadsheet programs, including Gnumeric.

== Installation
=== Standard Installation
   ruby test/ts_all.rb (optional)
   ruby install.rb

=== Gem Installation
   ruby test/ts_all.rb (optional)
   ruby spreadsheet-excel.gemspec
   gem install spreadsheet-excel-x.y.z.gem # where 'x.y.z' is the version

   or directly via RubyForge:
   gem install spreadsheet-excel
   
== Synopsis
   require "spreadsheet/excel"
   include Spreadsheet
   
   workbook = Excel.new("test.xls")
   
   format = Format.new
   format.color = "green"
   format.bold  = true
   
   worksheet = workbook.add_worksheet
   worksheet.write(0, 0, "Hello", format)
   worksheet.write(1, 1, ["Matz","Larry","Guido"])
   
   workbook.close

== What it doesn't do
   There is no support for formulas (yet).
   There is no support for worksheets greater than 7 MB.
   You cannot read/parse an existing spreadsheet with this package.

== Regarding formula support
   Simple formulas are easy enough, but to handle complex formulas in a
   reasonable fashion requires a parser.  John used "Parse::RecDescent" in his
   own code to parse formulas and I will need something similar to do so as
   well. Since I'm not too good at parsing, and I don't personally have the
   need for formula support, I'm more or less waiting for a patch.

== Regarding the 7MB limit
   Getting past the 7 MB limit requires an interface to the MS structured
   storage format.  This doesn't exist (yet) in Ruby.  For more on structured
   storage documents, download this: http://www.i3a.org/pdf/wg1n1017.pdf
   (there's a structured storage section).  That, and there is information
   about structured storage on the MSDN website at http://microsoft.msdn.com

== More information
   See the documentation in the 'doc' directory for more details.

== Known Bugs
   None that I'm aware of.  If you find any, please log them on the project
   page at http://rubyspreadsheet.sf.net.

== License
   Ruby's

== Copyright
   (C) 2005, Daniel J. Berger
   All Rights Reserved

== Author
   Daniel J. Berger
   djberg96 at gmail dot com
   IRC nickname: imperator/mok/rubyhacker1 (freenode)

== Maintainer 
   Hannes Wyss
   hannes.wyss@gmail.com
