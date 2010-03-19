#!/usr/bin/ruby -w

##############################################################################
#
# An example of adding document properties to a WriteExcel file.
#
# reverse('©'), August 2008, John McNamara, jmcnamara@cpan.org
#
# original written in Perl by John McNamara
# converted to Ruby by Hideo Nakamura, cxn03651@msj.biglobe.ne.jp
#
require 'rubygems'
require 'writeexcel'

workbook  = WriteExcel.new('properties.xls')
worksheet = workbook.add_worksheet
bp =1
workbook.set_properties(
    :title    => 'This is an example spreadsheet',
    :subject  => 'With document properties',
    :author   => 'Hideo NAKAMURA',
    :manager  => 'John McNamara',
    :company  => 'Rubygem',
    :category => 'Example spreadsheets',
    :keywords => 'Sample, Example, Properties',
    :comments => 'Created with Ruby and WriteExcel'
)


worksheet.set_column('A:A', 50)
worksheet.write('A1', 'Select File->Properties to see the file properties')

workbook.close
