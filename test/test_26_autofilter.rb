##########################################################################
# test_26_autofilter.rb
#
# Tests for the internal methods used to write the AUTOFILTER record.
#
# reverse('©'), September 2005, John McNamara, jmcnamara@cpan.org
#
#########################################################################
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

require "test/unit"
require 'writeexcel'

class TC_26_autofilter < Test::Unit::TestCase

  def test_26_autofilter
    @tests.each do |test|
      column     = test['column']
      expression = test['expression']
      tokens     = @worksheet.extract_filter_tokens(expression)
      tokens     = @worksheet.parse_filter_expression(expression, tokens)
  
      result     = @worksheet.store_autofilter(column, *tokens)
  
      target     = test['data'].join(" ")
  
      caption    = " \tfilter_column(#{column}, '#{expression}')"
  
      result     = unpack_record(result)
      assert_equal(target, result, caption)
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
    @workbook   = Spreadsheet::WriteExcel.new(@test_file)
    @worksheet  = @workbook.add_worksheet
    @tests = [
        {
            'column'        => 0,
            'expression'    => 'x =  Blanks',
            'data'          => [%w(
                                    9E 00 18 00 00 00 84 32 0C 02 00 00 00 00 00 00
                                    00 00 00 00 00 00 00 00 00 00 00 00
    
                               )]
        },
        {
            'column'        => 1,
            'expression'    => 'x =  Nonblanks',
            'data'          => [%w(
                                    9E 00 18 00 01 00 84 32 0E 05 00 00 00 00 00 00
                                    00 00 00 00 00 00 00 00 00 00 00 00
    
                               )]
        },
        {
            'column'        => 2,
            'expression'    => 'x >  1.001',
            'data'          => [%w(
                                    9E 00 18 00 02 00 80 32 04 04 6A BC 74 93 18 04
                                    F0 3F 00 00 00 00 00 00 00 00 00 00
    
                               )]
        },
        {
            'column'        => 3,
            'expression'    => 'x >= 1.001',
            'data'          => [%w(
                                      9E 00 18 00 03 00 80 32 04 06 6A BC 74 93 18 04
                                    F0 3F 00 00 00 00 00 00 00 00 00 00
    
                               )]
        },
        {
            'column'        => 4,
            'expression'    => 'x <  1.001',
            'data'          => [%w(
                                    9E 00 18 00 04 00 80 32 04 01 6A BC 74 93 18 04
                                    F0 3F 00 00 00 00 00 00 00 00 00 00
    
                               )]
        },
        {
            'column'        => 5,
            'expression'    => 'x <= 1.001',
            'data'          => [%w(
                                    9E 00 18 00 05 00 80 32 04 03 6A BC 74 93 18 04
                                    F0 3F 00 00 00 00 00 00 00 00 00 00
    
                               )]
        },
        {
            'column'        => 6,
            'expression'    => 'x >  1.001 and x <= 5.001',
            'data'          => [%w(
                                    9E 00 18 00 06 00 80 32 04 04 6A BC 74 93 18 04
                                    F0 3F 04 03 1B 2F DD 24 06 01 14 40
    
                               )]
        },
        {
            'column'        => 7,
            'expression'    => 'x >  1.001 or  x <= 5.001',
            'data'          => [%w(
                                    9E 00 18 00 07 00 81 32 04 04 6A BC 74 93 18 04
                                    F0 3F 04 03 1B 2F DD 24 06 01 14 40
    
                               )]
        },
        {
            'column'        => 8,
            'expression'    => 'x <> 2.001',
            'data'          => [%w(
                                    9E 00 18 00 08 00 80 32 04 05 35 5E BA 49 0C 02
                                    00 40 00 00 00 00 00 00 00 00 00 00
    
                               )]
        },
        {
            'column'        => 9,
            'expression'    => 'x =  1.001',
            'data'          => [%w(
                                    9E 00 1E 00 09 00 84 32 06 02 00 00 00 00 05 00
                                    00 00 00 00 00 00 00 00 00 00 00 00 00 31 2E 30
                                    30 31
    
                               )]
        },
        {
            'column'        => 10,
            'expression'    => 'x =  West',
            'data'          => [%w(
                                    9E 00 1D 00 0A 00 84 32 06 02 00 00 00 00 04 00
                                    00 00 00 00 00 00 00 00 00 00 00 00 00 57 65 73
                                    74
    
                               )]
        },
        {
            'column'        => 11,
            'expression'    => 'x =  East',
            'data'          => [%w(
                                    9E 00 1D 00 0B 00 84 32 06 02 00 00 00 00 04 00
                                    00 00 00 00 00 00 00 00 00 00 00 00 00 45 61 73
                                    74
    
                               )]
        },
        {
            'column'        => 12,
            'expression'    => 'x <> West',
            'data'          => [%w(
                                    9E 00 1D 00 0C 00 80 32 06 05 00 00 00 00 04 00
                                    00 00 00 00 00 00 00 00 00 00 00 00 00 57 65 73
                                    74
    
                               )]
        },
        {
            'column'        => 13,
            'expression'    => 'x =~ b*',
            'data'          => [%w(
                                    9E 00 1B 00 0D 00 80 32 06 02 00 00 00 00 02 00
                                    00 00 00 00 00 00 00 00 00 00 00 00 00 62 2A
    
                               )]
        },
        {
            'column'        => 14,
            'expression'    => 'x !~ b*',
            'data'          => [%w(
                                    9E 00 1B 00 0E 00 80 32 06 05 00 00 00 00 02 00
                                    00 00 00 00 00 00 00 00 00 00 00 00 00 62 2A
    
                               )]
        },
        {
            'column'        => 15,
            'expression'    => 'x =~ *b',
            'data'          => [%w(
                                    9E 00 1B 00 0F 00 80 32 06 02 00 00 00 00 02 00
                                    00 00 00 00 00 00 00 00 00 00 00 00 00 2A 62
    
                               )]
        },
        {
            'column'        => 16,
            'expression'    => 'x !~ *b',
            'data'          => [%w(
                                    9E 00 1B 00 10 00 80 32 06 05 00 00 00 00 02 00
                                    00 00 00 00 00 00 00 00 00 00 00 00 00 2A 62
    
                               )]
        },
        {
            'column'        => 17,
            'expression'    => 'x =~ *b*',
            'data'          => [%w(
                                    9E 00 1C 00 11 00 80 32 06 02 00 00 00 00 03 00
                                    00 00 00 00 00 00 00 00 00 00 00 00 00 2A 62 2A
    
                               )]
        },
        {
            'column'        => 18,
            'expression'    => 'x !~ *b*',
            'data'          => [%w(
                                    9E 00 1C 00 12 00 80 32 06 05 00 00 00 00 03 00
                                    00 00 00 00 00 00 00 00 00 00 00 00 00 2A 62 2A
    
                               )]
        },
        {
            'column'        => 19,
            'expression'    => 'x =  fo?',
            'data'          => [%w(
                                    9E 00 1C 00 13 00 80 32 06 02 00 00 00 00 03 00
                                    00 00 00 00 00 00 00 00 00 00 00 00 00 66 6F 3F
    
                               )]
        },
        {
            'column'        => 20,
            'expression'    => 'x =  fo~?',
            'data'          => [%w(
                                    9E 00 1D 00 14 00 80 32 06 02 00 00 00 00 04 00
                                    00 00 00 00 00 00 00 00 00 00 00 00 00 66 6F 7E
                                    3F
    
                               )]
        },
        {
            'column'        => 21,
            'expression'    => 'x =  East and x = West',
            'data'          => [%w(
                                    9E 00 22 00 15 00 8C 32 06 02 00 00 00 00 04 00
                                    00 00 06 02 00 00 00 00 04 00 00 00 00 45 61 73
                                    74 00 57 65 73 74
    
                               )]
        },
        {
            'column'        => 22,
            'expression'    => 'top 10 items',
            'data'          => [%w(
                                    9E 00 18 00 16 00 30 05 04 06 00 00 00 00 00 00
                                    00 00 00 00 00 00 00 00 00 00 00 00
    
                               )]
        },
        {
            'column'        => 23,
            'expression'    => 'top 10 %',
            'data'          => [%w(
                                    9E 00 18 00 17 00 70 05 04 06 00 00 00 00 00 00
                                    00 00 00 00 00 00 00 00 00 00 00 00
    
                               )]
        },
        {
            'column'        => 24,
            'expression'    => 'bottom 10 items',
            'data'          => [%w(
                                    9E 00 18 00 18 00 10 05 04 03 00 00 00 00 00 00
                                    00 00 00 00 00 00 00 00 00 00 00 00
    
                               )]
        },
        {
            'column'        => 25,
            'expression'    => 'bottom 10 %',
            'data'          => [%w(
                                    9E 00 18 00 19 00 50 05 04 03 00 00 00 00 00 00
                                    00 00 00 00 00 00 00 00 00 00 00 00
    
                               )]
        },
        {
            'column'        => 26,
            'expression'    => 'top 5 items',
            'data'          => [%w(
                                    9E 00 18 00 1A 00 B0 02 04 06 00 00 00 00 00 00
                                    00 00 00 00 00 00 00 00 00 00 00 00
    
                               )]
        },
        {
            'column'        => 27,
            'expression'    => 'top 100 items',
            'data'          => [%w(
                                    9E 00 18 00 1B 00 30 32 04 06 00 00 00 00 00 00
                                    00 00 00 00 00 00 00 00 00 00 00 00
    
                               )]
        },
        {
            'column'        => 28,
            'expression'    => 'top 101 items',
            'data'          => [%w(
                                    9E 00 18 00 1C 00 B0 32 04 06 00 00 00 00 00 00
                                    00 00 00 00 00 00 00 00 00 00 00 00
    
                               )]
        }
      ]
  end

  def teardown
    @workbook.close
    File.unlink(@test_file) if FileTest.exist?(@test_file)
  end

end
