#######################################################
# tc_example_match.rb
#
# Test suite for matching with xls file made by perl.
#######################################################
base = File.basename(Dir.pwd)
if base == "test" || base =~ /spreadsheet/i
  Dir.chdir("..") if base == "test"
  $LOAD_PATH.unshift(Dir.pwd + "/lib/spreadsheet")
  Dir.chdir("test") rescue nil
end

require "test/unit"
require "excel"
include Spreadsheet

class TC_example_match < Test::Unit::TestCase

  def setup
    @filename  = "tc_example_match.xls"
    @filename2 = "tc_example_match2.xls"
  end

  def teardown
    File.delete(@filename)  if File.exist?(@filename)
    File.delete(@filename2) if File.exist?(@filename2)
  end

  def test_a_simple
    workbook  = Excel.new(@filename);
    worksheet = workbook.add_worksheet
    
    # The general syntax is write(row, column, token). Note that row and
    # column are zero indexed
    #
    
    # Write some text
    worksheet.write(0, 0,  "Hi Excel!")
    
    
    # Write some numbers
    worksheet.write(2, 0,  3)          # Writes 3
    worksheet.write(3, 0,  3.00000)    # Writes 3
    worksheet.write(4, 0,  3.00001)    # Writes 3.00001
    worksheet.write(5, 0,  3.14159)    # TeX revision no.?
    
    
    # Write some formulas
    worksheet.write(7, 0,  '=A3 + A6')
    worksheet.write(8, 0,  '=IF(A5>3,"Yes", "No")')
    
    
    # Write a hyperlink
    worksheet.write(10, 0, 'http://www.perl.com/')
    
    # File save
    workbook.close
    
    # do assertion
    compare_file("perl_output/a_simple.xls", @filename)
  end

  def test_regions
    workbook = Excel.new(@filename)

    # Add some worksheets
    north = workbook.add_worksheet("North")
    south = workbook.add_worksheet("South")
    east  = workbook.add_worksheet("East")
    west  = workbook.add_worksheet("West")
    
    # Add a Format
    format = workbook.add_format()
    format.set_bold()
    format.set_color('blue')
    
    # Add a caption to each worksheet
    workbook.sheets.each do |worksheet|
        worksheet.write(0, 0, "Sales", format)
    end
    
    # Write some data
    north.write(0, 1, 200000)
    south.write(0, 1, 100000)
    east.write(0, 1, 150000)
    west.write(0, 1, 100000)
    
    # Set the active worksheet
    bp=1
    south.activate()
    
    # Set the width of the first column
    south.set_column(0, 0, 20)
    
    # Set the active cell
    south.set_selection(0, 1)
    
    workbook.close

    # do assertion
    compare_file("perl_output/regions.xls", @filename)
  end

  def test_stats
    workbook = Excel.new(@filename)
    worksheet = workbook.add_worksheet('Test data')

    # Set the column width for columns 1
    worksheet.set_column(0, 0, 20)

    # Create a format for the headings
    format = workbook.add_format
    format.set_bold

    # Write the sample data
    worksheet.write(0, 0, 'Sample', format)
    worksheet.write(0, 1, 1)
    worksheet.write(0, 2, 2)
    worksheet.write(0, 3, 3)
    worksheet.write(0, 4, 4)
    worksheet.write(0, 5, 5)
    worksheet.write(0, 6, 6)
    worksheet.write(0, 7, 7)
    worksheet.write(0, 8, 8)

    worksheet.write(1, 0, 'Length', format)
    worksheet.write(1, 1, 25.4)
    worksheet.write(1, 2, 25.4)
    worksheet.write(1, 3, 24.8)
    worksheet.write(1, 4, 25.0)
    worksheet.write(1, 5, 25.3)
    worksheet.write(1, 6, 24.9)
    worksheet.write(1, 7, 25.2)
    worksheet.write(1, 8, 24.8)

    # Write some statistical functions
    worksheet.write(4,  0, 'Count', format)
    worksheet.write(4,  1, '=COUNT(B1:I1)')

    worksheet.write(5,  0, 'Sum', format)
    worksheet.write(5,  1, '=SUM(B2:I2)')

    worksheet.write(6,  0, 'Average', format)
    worksheet.write(6,  1, '=AVERAGE(B2:I2)')

    worksheet.write(7,  0, 'Min', format)
    worksheet.write(7,  1, '=MIN(B2:I2)')

    worksheet.write(8,  0, 'Max', format)
    worksheet.write(8,  1, '=MAX(B2:I2)')

    worksheet.write(9,  0, 'Standard Deviation', format)
    worksheet.write(9,  1, '=STDEV(B2:I2)')
    
    worksheet.write(10, 0, 'Kurtosis', format)
    worksheet.write(10, 1, '=KURT(B2:I2)')

    workbook.close

    # do assertion
    compare_file("perl_output/regions.xls", @filename)
  end

  def test_hidden
    workbook   = Excel.new(@filename)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet

    # Sheet2 won't be visible until it is unhidden in Excel.
    worksheet2.hide

    worksheet1.write(0, 0, 'Sheet2 is hidden')
    worksheet2.write(0, 0, 'How did you find me?')
    worksheet3.write(0, 0, 'Sheet2 is hidden')

    workbook.close

    # do assertion
    compare_file("perl_output/hidden.xls", @filename)
  end

  def test_hyperlink1
    # Create a new workbook and add a worksheet
    workbook  = Excel.new(@filename)
    worksheet = workbook.add_worksheet('Hyperlinks')
    
    # Format the first column
    worksheet.set_column('A:A', 30)
    worksheet.set_selection('B1')
    
    
    # Add a sample format
    format = workbook.add_format
    format.set_size(12)
    format.set_bold
    format.set_color('red')
    format.set_underline
    
    
    # Write some hyperlinks
    worksheet.write('A1', 'http://www.perl.com/'                )
    worksheet.write('A3', 'http://www.perl.com/', 'Perl home'   )
    worksheet.write('A5', 'http://www.perl.com/', nil, format)
    worksheet.write('A7', 'mailto:jmcnamara@cpan.org', 'Mail me')
    
    # Write a URL that isn't a hyperlink
    worksheet.write_string('A9', 'http://www.perl.com/')
    
    workbook.close

    # do assertion
    compare_file("perl_output/hyperlink.xls", @filename)
  end

  def test_copyformat
    # Create workbook1
    workbook1       = Excel.new(@filename)
    worksheet1      = workbook1.add_worksheet
    format1a        = workbook1.add_format
    format1b        = workbook1.add_format
    
    # Create workbook2
    workbook2       = Excel.new(@filename2)
    worksheet2      = workbook2.add_worksheet
    format2a        = workbook2.add_format
    format2b        = workbook2.add_format
    
    # Create a global format object that isn't tied to a workbook
    global_format   = Format.new
    
    # Set the formatting
    global_format.set_color('blue')
    global_format.set_bold
    global_format.set_italic
    
    # Create another example format
    format1b.set_color('red')
    
    # Copy the global format properties to the worksheet formats
    format1a.copy(global_format)
    format2a.copy(global_format)
    
    # Copy a format from worksheet1 to worksheet2
    format2b.copy(format1b)
    
    # Write some output
    worksheet1.write(0, 0, "Ciao", format1a)
    worksheet1.write(1, 0, "Ciao", format1b)
    
    worksheet2.write(0, 0, "Hello", format2a)
    worksheet2.write(1, 0, "Hello", format2b)
    workbook1.close
    workbook2.close

    # do assertion
    compare_file("perl_output/workbook1.xls", @filename)
    compare_file("perl_output/workbook2.xls", @filename2)
  end

  def test_data_validate
    workbook  = Excel.new(@filename)
    worksheet = workbook.add_worksheet
    
    # Add a format for the header cells.
    header_format = workbook.add_format(
                                                :border      => 1,
                                                :bg_color    => 43,
                                                :bold        => 1,
                                                :text_wrap   => 1,
                                                :valign      => 'vcenter',
                                                :indent      => 1
                                             )
    
    # Set up layout of the worksheet.
    worksheet.set_column('A:A', 64)
    worksheet.set_column('B:B', 15)
    worksheet.set_column('D:D', 15)
    worksheet.set_row(0, 36)
    worksheet.set_selection('B3')
    
    
    # Write the header cells and some data that will be used in the examples.
    row = 0
    heading1 = 'Some examples of data validation in Spreadsheet::WriteExcel'
    heading2 = 'Enter values in this column'
    heading3 = 'Sample Data'
    
    worksheet.write('A1', heading1, header_format)
    worksheet.write('B1', heading2, header_format)
    worksheet.write('D1', heading3, header_format)
    
    worksheet.write('D3', ['Integers',   1, 10])
    worksheet.write('D4', ['List data', 'open', 'high', 'close'])
    worksheet.write('D5', ['Formula',   '=AND(F5=50,G5=60)', 50, 60])
    
    
    #
    # Example 1. Limiting input to an integer in a fixed range.
    #
    txt = 'Enter an integer between 1 and 10'
    row += 2
    
    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
        {
            :validate        => 'integer',
            :criteria        => 'between',
            :minimum         => 1,
            :maximum         => 10
        })
    
    
    #
    # Example 2. Limiting input to an integer outside a fixed range.
    #
    txt = 'Enter an integer that is not between 1 and 10 (using cell references)'
    row += 2
    
    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
        {
            :validate        => 'integer',
            :criteria        => 'not between',
            :minimum         => '=E3',
            :maximum         => '=F3'
        })
    
    
    #
    # Example 3. Limiting input to an integer greater than a fixed value.
    #
    txt = 'Enter an integer greater than 0'
    row += 2
    
    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
        {
            :validate        => 'integer',
            :criteria        => '>',
            :value           => 0
        })
    
    
    #
    # Example 4. Limiting input to an integer less than a fixed value.
    #
    txt = 'Enter an integer less than 10'
    row += 2
    
    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
        {
            :validate        => 'integer',
            :criteria        => '<',
            :value           => 10
        })
    
    
    #
    # Example 5. Limiting input to a decimal in a fixed range.
    #
    txt = 'Enter a decimal between 0.1 and 0.5'
    row += 2
    
    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
        {
            :validate        => 'decimal',
            :criteria        => 'between',
            :minimum         => 0.1,
            :maximum         => 0.5
        })
    
    
    #
    # Example 6. Limiting input to a value in a dropdown list.
    #
    txt = 'Select a value from a drop down list'
    row += 2
    bp=1
    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
        {
            :validate        => 'list',
            :source          => ['open', 'high', 'close']
        })
    
    
    #
    # Example 6. Limiting input to a value in a dropdown list.
    #
    txt = 'Select a value from a drop down list (using a cell range)'
    row += 2
    
    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
        {
            :validate        => 'list',
            :source          => '=E4:G4'
        })
    
    
    #
    # Example 7. Limiting input to a date in a fixed range.
    #
    txt = 'Enter a date between 1/1/2008 and 12/12/2008'
    row += 2
    
    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
        {
            :validate        => 'date',
            :criteria        => 'between',
            :minimum         => '2008-01-01T',
            :maximum         => '2008-12-12T'
        })
    
    
    #
    # Example 8. Limiting input to a time in a fixed range.
    #
    txt = 'Enter a time between 6:00 and 12:00'
    row += 2
    
    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
        {
            :validate        => 'time',
            :criteria        => 'between',
            :minimum         => 'T06:00',
            :maximum         => 'T12:00'
        })
    
    
    #
    # Example 9. Limiting input to a string greater than a fixed length.
    #
    txt = 'Enter a string longer than 3 characters'
    row += 2
    
    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
        {
            :validate        => 'length',
            :criteria        => '>',
            :value           => 3
        })
    
    
    #
    # Example 10. Limiting input based on a formula.
    #
    txt = 'Enter a value if the following is true "=AND(F5=50,G5=60)"'
    row += 2
    
    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
        {
            :validate        => 'custom',
            :value           => '=AND(F5=50,G5=60)'
        })
    
    
    #
    # Example 11. Displaying and modify data validation messages.
    #
    txt = 'Displays a message when you select the cell'
    row += 2
    
    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
        {
            :validate      => 'integer',
            :criteria      => 'between',
            :minimum       => 1,
            :maximum       => 100,
            :input_title   => 'Enter an integer:',
            :input_message => 'between 1 and 100'
        })
    
    
    #
    # Example 12. Displaying and modify data validation messages.
    #
    txt = 'Display a custom error message when integer isn\'t between 1 and 100'
    row += 2
    
    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
        {
            :validate      => 'integer',
            :criteria      => 'between',
            :minimum       => 1,
            :maximum       => 100,
            :input_title   => 'Enter an integer:',
            :input_message => 'between 1 and 100',
            :error_title   => 'Input value is not valid!',
            :error_message => 'It should be an integer between 1 and 100'
        })
    
    
    #
    # Example 13. Displaying and modify data validation messages.
    #
    txt = 'Display a custom information message when integer isn\'t between 1 and 100'
    row += 2
    
    worksheet.write(row, 0, txt)
    worksheet.data_validation(row, 1,
        {
            :validate      => 'integer',
            :criteria      => 'between',
            :minimum       => 1,
            :maximum       => 100,
            :input_title   => 'Enter an integer:',
            :input_message => 'between 1 and 100',
            :error_title   => 'Input value is not valid!',
            :error_message => 'It should be an integer between 1 and 100',
            :error_type    => 'information'
        })
    
    workbook.close

    # do assertion
    compare_file("perl_output/data_validate.xls", @filename)
  end

  def test_merge1
    workbook  = Excel.new(@filename)
    worksheet = workbook.add_worksheet

    # Increase the cell size of the merged cells to highlight the formatting.
    worksheet.set_column('B:D', 20)
    worksheet.set_row(2, 30)

    # Create a merge format
    format = workbook.add_format(:center_across => 1)

    # Only one cell should contain text, the others should be blank.
    worksheet.write(2, 1, "Center across selection", format)
    worksheet.write_blank(2, 2,                 format)
    worksheet.write_blank(2, 3,                 format)

    workbook.close

    # do assertion
    compare_file("perl_output/merge1.xls", @filename)
  end

  def test_merge2
    workbook  = Excel.new(@filename)
    worksheet = workbook.add_worksheet

    # Increase the cell size of the merged cells to highlight the formatting.
    worksheet.set_column(1, 2, 30)
    worksheet.set_row(2, 40)

    # Create a merged format
    format = workbook.add_format(
                                        :center_across   => 1,
                                        :bold            => 1,
                                        :size            => 15,
                                        :pattern         => 1,
                                        :border          => 6,
                                        :color           => 'white',
                                        :fg_color        => 'green',
                                        :border_color    => 'yellow',
                                        :align           => 'vcenter'
                                  )

    # Only one cell should contain text, the others should be blank.
    worksheet.write(2, 1, "Center across selection", format)
    worksheet.write_blank(2, 2,                      format)
    workbook.close

    # do assertion
    compare_file("perl_output/merge2.xls", @filename)
  end
  
  def test_merge3
    workbook  = Excel.new(@filename)
    worksheet = workbook.add_worksheet()
    
    # Increase the cell size of the merged cells to highlight the formatting.
    [1, 3,6,7].each { |row| worksheet.set_row(row, 30) }
    worksheet.set_column('B:D', 20)
    
    ###############################################################################
    #
    # Example 1: Merge cells containing a hyperlink using write_url_range()
    # and the standard Excel 5+ merge property.
    #
    format1 = workbook.add_format(
                                        :center_across   => 1,
                                        :border          => 1,
                                        :underline       => 1,
                                        :color           => 'blue'
                                 )
    
    # Write the cells to be merged
    worksheet.write_url_range('B2:D2', 'http://www.perl.com', format1)
    worksheet.write_blank('C2', format1)
    worksheet.write_blank('D2', format1)
    
    
    
    ###############################################################################
    #
    # Example 2: Merge cells containing a hyperlink using merge_range().
    #
    format2 = workbook.add_format(
                                        :border      => 1,
                                        :underline   => 1,
                                        :color       => 'blue',
                                        :align       => 'center',
                                        :valign      => 'vcenter'
                                 )
    
    # Merge 3 cells
    worksheet.merge_range('B4:D4', 'http://www.perl.com', format2)
    
    
    # Merge 3 cells over two rows
    worksheet.merge_range('B7:D8', 'http://www.perl.com', format2)
    
    workbook.close

    # do assertion
    compare_file("perl_output/merge3.xls", @filename)
  end

  def test_merge4
    # Create a new workbook and add a worksheet
    workbook  = Excel.new(@filename)
    worksheet = workbook.add_worksheet
    
    # Increase the cell size of the merged cells to highlight the formatting.
    (1..11).each { |row| worksheet.set_row(row, 30) }
    worksheet.set_column('B:D', 20)
    
    ###############################################################################
    #
    # Example 1: Text centered vertically and horizontally
    #
    format1 = workbook.add_format(
                                        :border  => 6,
                                        :bold    => 1,
                                        :color   => 'red',
                                        :valign  => 'vcenter',
                                        :align   => 'center'
                                       )
    
    worksheet.merge_range('B2:D3', 'Vertical and horizontal', format1)
    
    
    ###############################################################################
    #
    # Example 2: Text aligned to the top and left
    #
    format2 = workbook.add_format(
                                        :border  => 6,
                                        :bold    => 1,
                                        :color   => 'red',
                                        :valign  => 'top',
                                        :align   => 'left'
                                      )
    
    worksheet.merge_range('B5:D6', 'Aligned to the top and left', format2)
    
    ###############################################################################
    #
    # Example 3:  Text aligned to the bottom and right
    #
    format3 = workbook.add_format(
                                        :border  => 6,
                                        :bold    => 1,
                                        :color   => 'red',
                                        :valign  => 'bottom',
                                        :align   => 'right'
                                      )
    
    worksheet.merge_range('B8:D9', 'Aligned to the bottom and right', format3)
    
    ###############################################################################
    #
    # Example 4:  Text justified (i.e. wrapped) in the cell
    #
    format4 = workbook.add_format(
                                        :border  => 6,
                                        :bold    => 1,
                                        :color   => 'red',
                                        :valign  => 'top',
                                        :align   => 'justify'
                                      )
    
    worksheet.merge_range('B11:D12', 'Justified: '+'so on and '*18, format4)
    
    workbook.close

    # do assertion
    compare_file("perl_output/merge3.xls", @filename)
  end

  def test_merge5
    # Create a new workbook and add a worksheet
    workbook  = Excel.new(@filename)
    worksheet = workbook.add_worksheet
    
    
    # Increase the cell size of the merged cells to highlight the formatting.
    (3..8).each { |col| worksheet.set_row(col, 36) }
    [1, 3, 5].each { |n| worksheet.set_column(n, n, 15) }
    
    
    ###############################################################################
    #
    # Rotation 1, letters run from top to bottom
    #
    format1 = workbook.add_format(
                                        :border      => 6,
                                        :bold        => 1,
                                        :color       => 'red',
                                        :valign      => 'vcentre',
                                        :align       => 'centre',
                                        :rotation    => 270
                                      )
    
    
    worksheet.merge_range('B4:B9', 'Rotation 270', format1)
    
    
    ###############################################################################
    #
    # Rotation 2, 90° anticlockwise
    #
    format2 = workbook.add_format(
                                        :border      => 6,
                                        :bold        => 1,
                                        :color       => 'red',
                                        :valign      => 'vcentre',
                                        :align       => 'centre',
                                        :rotation    => 90
                                      )
    
    
    worksheet.merge_range('D4:D9', 'Rotation 90°', format2)
    
    
    
    ###############################################################################
    #
    # Rotation 3, 90° clockwise
    #
    format3 = workbook.add_format(
                                        :border      => 6,
                                        :bold        => 1,
                                        :color       => 'red',
                                        :valign      => 'vcentre',
                                        :align       => 'centre',
                                        :rotation    => -90
                                      )
    
    
    worksheet.merge_range('F4:F9', 'Rotation -90°', format3)
    
    workbook.close
    
    # do assertion
    compare_file("perl_output/merge3.xls", @filename)
  end

  def compare_file(expected, target)
    fh_e = File.open(expected, "r")
    fh_t = File.open(target, "r")
    while true do
      e1 = fh_e.read(1)
      t1 = fh_t.read(1)
      if e1.nil?
        assert( t1.nil?, "#{expexted} is EOF but #{target} is NOT EOF.")
        break
      elsif t1.nil?
        assert( e1.nil?, '#{target} is EOF but #{expected} is NOT EOF.')
        break
      end
      assert_equal(e1, t1, sprintf(" #{expected} = '%s' but #{target} = '%s'", e1, t1))
      break
    end
    fh_e.close
    fh_t.close
  end


end
