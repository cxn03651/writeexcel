##########################################################################
# test_42_set_properties.rb
#
# Tests for Workbook property_sets() interface.
#
# reverse('©'), September 2005, John McNamara, jmcnamara@cpan.org
#
#########################################################################
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

require 'test/unit'
require 'writeexcel'

class TC_set_properties < Test::Unit::TestCase

  def test_same_as_previous_plus_creation_date
    test_file = 'temp_test_file.xls'
    smiley = '☺'   # chr 0x263A;    in perl
    
    workbook  = Spreadsheet::WriteExcel.new(test_file)
    worksheet = workbook.add_worksheet

    ###############################################################################
    #
    # Test 1. _get_property_set_codepage() for default latin1 strings.
    #
    params ={
                    :title       => 'Title',
                    :subject     => 'Subject',
                    :author      => 'Author',
                    :keywords    => 'Keywords',
                    :comments    => 'Comments',
                    :last_author => 'Username',
            }

    strings = %w(title subject author keywords comments last_author)

    caption    = " \t_get_property_set_codepage('latin1')"
    target     = 0x04E4
    
    result     = workbook.get_property_set_codepage(params, strings)
    assert_equal(target, result, caption)

    ###############################################################################
    #
    # Test 2. _get_property_set_codepage() for manual utf8 strings.
    #
    
    params =   {
                    :title       => 'Title',
                    :subject     => 'Subject',
                    :author      => 'Author',
                    :keywords    => 'Keywords',
                    :comments    => 'Comments',
                    :last_author => 'Username',
                    :utf8        => 1,
            }
    
    strings = %w(title subject author keywords comments last_author)

    caption    = " \t_get_property_set_codepage('utf8')"
    target     = 0xFDE9

    result     = workbook.get_property_set_codepage(params, strings)
    assert_equal(target, result, caption)

    ###############################################################################
    #
    # Test 3. _get_property_set_codepage() for perl 5.8 utf8 strings.
    #
    params =   {
                    :title       => 'Title' + smiley,
                    :subject     => 'Subject',
                    :author      => 'Author',
                    :keywords    => 'Keywords',
                    :comments    => 'Comments',
                    :last_author => 'Username',
                }
    
    strings = %w(title subject author keywords comments last_author)

    caption    = " \t_get_property_set_codepage('utf8')";
    target     = 0xFDE9;
    
    result     = workbook.get_property_set_codepage(params, strings)
    assert_equal(target, result, caption)

  end

  ###############################################################################
  #
  # Unpack the binary data into a format suitable for printing in tests.
  #
  def unpack_record(data)
    data.unpack('C*').map! {|c| sprintf("%02X", c) }.join(' ')
  end

end
