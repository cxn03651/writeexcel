##########################################################################
# test_28_autofilter.rb
#
# Tests for the token parsing methods used to parse autofilter expressions.
#
# reverse('©'), September 2005, John McNamara, jmcnamara@cpan.org
#
# original written in Perl by John McNamara
# converted to Ruby by Hideo Nakamura, cxn03651@msj.biglobe.ne.jp
#
#########################################################################
require 'helper'

class TC_28_autofilter < Test::Unit::TestCase

  def test_28_autofilter
    @tests.each do |test|
      expression = test[0]
      expected   = test[1]
      tokens     = @worksheet.extract_filter_tokens(expression)
      result     = @worksheet.parse_filter_expression(expression, tokens)

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
    t = Time.now.strftime("%Y%m%d")
    path = "temp#{t}-#{$$}-#{rand(0x100000000).to_s(36)}"
    @test_file           = File.join(Dir.tmpdir, path)
    @workbook   = WriteExcel.new(@test_file)
    @worksheet  = @workbook.add_worksheet
    @tests = [
    [
        'x =  2000',
        [2, 2000],
    ],

    [
        'x == 2000',
        [2, 2000],
    ],

    [
        'x =~ 2000',
        [2, 2000],
    ],

    [
        'x eq 2000',
        [2, 2000],
    ],

    [
        'x <> 2000',
        [5, 2000],
    ],

    [
        'x != 2000',
        [5, 2000],
    ],

    [
        'x ne 2000',
        [5, 2000],
    ],

    [
        'x !~ 2000',
        [5, 2000],
    ],

    [
        'x >  2000',
        [4, 2000],
    ],

    [
        'x <  2000',
        [1, 2000],
    ],

    [
        'x >= 2000',
        [6, 2000],
    ],

    [
        'x <= 2000',
        [3, 2000],
    ],

    [
        'x >  2000 and x <  5000',
        [4,  2000, 0, 1, 5000],
    ],

    [
        'x >  2000 &&  x <  5000',
        [4,  2000, 0, 1, 5000],
    ],

    [
        'x >  2000 or  x <  5000',
        [4,  2000, 1, 1, 5000],
    ],

    [
        'x >  2000 ||  x <  5000',
        [4,  2000, 1, 1, 5000],
    ],

    [
        'x =  Blanks',
        [2, 'blanks'],
    ],

    [
        'x =  NonBlanks',
        [2, 'nonblanks'],
    ],

    [
        'x <> Blanks',
        [2, 'nonblanks'],
    ],

    [
        'x <> NonBlanks',
        [2, 'blanks'],
    ],

    [
        'Top 10 Items',
        [30, 10],
    ],

    [
        'Top 20 %',
        [31, 20],
    ],

    [
        'Bottom 5 Items',
        [32, 5],
    ],

    [
        'Bottom 101 %',
        [33, 101],
    ],
      ]
  end

  def teardown
    @workbook.close
    File.unlink(@test_file) if FileTest.exist?(@test_file)
  end

end
