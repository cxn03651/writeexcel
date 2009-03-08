##########################################################################
# test_27_autofilter.rb
#
# Tests for the token extraction method used to parse autofilter expressions.
#
# reverse('©'), September 2005, John McNamara, jmcnamara@cpan.org
#
#########################################################################
base = File.basename(Dir.pwd)
if base == "test" || base =~ /spreadsheet/i
  Dir.chdir("..") if base == "test"
  $LOAD_PATH.unshift(Dir.pwd + "/lib/spreadsheet")
  Dir.chdir("test") rescue nil
end

require "test/unit"
require "biffwriter"
require "olewriter"
require "format"
require "formula"
require "worksheet"
require "workbook"
require "excel"
include Spreadsheet


class TC_27_autofilter < Test::Unit::TestCase

  def test_27_autofilter
    @tests.each do |test|
      expression = test[0]
      expected   = test[1]
      result     = @worksheet.extract_filter_tokens(expression)
  
      testname   = expression || 'none'
      
      assert_equal(expected, result, testname)
    end
  end

  ###############################################################################
  #
  # Unpack the binary data into a format suitable for printing in tests.
  #
  def unpack_record(data)
    data.unpack('C*').map! {|c| sprintf("%02X", c) }.join(' ')
  end

  def setup
    @test_file  = 'temp_test_file.xls'
    @workbook   = Excel.new(@test_file)
    @worksheet  = @workbook.add_worksheet
    @tests = [
    [
        nil,
        [],
    ],

    [
        '',
        [],
    ],

    [
        '0 <  2000',
        [0, '<', 2000],
    ],

    [
        'x <  2000',
        ['x', '<', 2000],
    ],

    [
        'x >  2000',
        ['x', '>', 2000],
    ],

    [
        'x == 2000',
        ['x', '==', 2000],
    ],

    [
        'x >  2000 and x <  5000',
        ['x', '>',  2000, 'and', 'x', '<', 5000],
    ],

    [
        'x = "foo"',
        ['x', '=', 'foo'],
    ],

    [
        'x = foo',
        ['x', '=', 'foo'],
    ],

    [
        'x = "foo bar"',
        ['x', '=', 'foo bar'],
    ],

    [
        'x = "foo "" bar"',
        ['x', '=', 'foo " bar'],
    ],

    [
        'x = "foo bar" or x = "bar foo"',
        ['x', '=', 'foo bar', 'or', 'x', '=', 'bar foo'],
    ],

    [
        'x = "foo "" bar" or x = "bar "" foo"',
        ['x', '=', 'foo " bar', 'or', 'x', '=', 'bar " foo'],
    ],

    [
        'x = """"""""',
        ['x', '=', '"""'],
    ],

    [
        'x = Blanks',
        ['x', '=', 'Blanks'],
    ],

    [
        'x = NonBlanks',
        ['x', '=', 'NonBlanks'],
    ],

    [
        'top 10 %',
        ['top', 10, '%'],
    ],

    [
        'top 10 items',
        ['top', 10, 'items'],
    ],

      ]
  end

end
