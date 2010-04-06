require 'helper'
require 'stringio'

class TC_example_match < Test::Unit::TestCase

  TEST_DIR    = File.expand_path(File.dirname(__FILE__))
  PERL_OUTDIR = File.join(TEST_DIR, 'perl_output')

  def setup
    @file  = StringIO.new
=begin
    t = Time.now.strftime("%Y%m%d")
    path = "temp#{t}-#{$$}-#{rand(0x100000000).to_s(36)}"
    @test_file  = File.join(Dir.tmpdir, path)
    @filename  = @test_file
    @filename2 = @test_file + "2"
=end
  end

  def test_a_simple
    workbook  = WriteExcel.new(@file)
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
    compare_file("#{PERL_OUTDIR}/a_simple.xls", @file)
  end

  def test_autofilter
    workbook = WriteExcel.new(@file)

    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet
    worksheet4 = workbook.add_worksheet
    worksheet5 = workbook.add_worksheet
    worksheet6 = workbook.add_worksheet

    bold       = workbook.add_format(:bold => 1)

    # Extract the data embedded at the end of this file.
    headings = %w(Region    Item      Volume    Month)
    data = get_data_for_autofilter

    # Set up several sheets with the same data.
    workbook.sheets.each do |worksheet|
        worksheet.set_column('A:D', 12)
        worksheet.set_row(0, 20, bold)
        worksheet.write('A1', headings)
    end

    ###############################################################################
    #
    # Example 1. Autofilter without conditions.
    #

    worksheet1.autofilter('A1:D51')
    worksheet1.write('A2', [data])

    ###############################################################################
    #
    #
    # Example 2. Autofilter with a filter condition in the first column.
    #

    # The range in this example is the same as above but in row-column notation.
    worksheet2.autofilter(0, 0, 50, 3)

    # The placeholder "Region" in the filter is ignored and can be any string
    # that adds clarity to the expression.
    #
    worksheet2.filter_column(0, 'Region eq East')

    #
    # Hide the rows that don't match the filter criteria.
    #
    row = 1

    data.each do |row_data|
        region = row_data[0]

        if region == 'East'
            # Row is visible.
        else
            # Hide row.
            worksheet2.set_row(row, nil, nil, 1)
        end

        worksheet2.write(row, 0, row_data)
        row += 1
    end


    ###############################################################################
    #
    #
    # Example 3. Autofilter with a dual filter condition in one of the columns.
    #

    worksheet3.autofilter('A1:D51')

    worksheet3.filter_column('A', 'x eq East or x eq South')

    #
    # Hide the rows that don't match the filter criteria.
    #
    row = 1

    data.each do |row_data|
        region = row_data[0]

        if region == 'East' || region == 'South'
            # Row is visible.
        else
            # Hide row.
            worksheet3.set_row(row, nil, nil, 1)
        end

        worksheet3.write(row, 0, row_data)
        row += 1
    end


    ###############################################################################
    #
    #
    # Example 4. Autofilter with filter conditions in two columns.
    #

    worksheet4.autofilter('A1:D51')

    worksheet4.filter_column('A', 'x eq East')
    worksheet4.filter_column('C', 'x > 3000 and x < 8000' )

    #
    # Hide the rows that don't match the filter criteria.
    #
    row = 1

    data.each do |row_data|
        region = row_data[0]
        volume = row_data[2]

        if region == 'East' && volume >  3000   && volume < 8000
            # Row is visible.
        else
            # Hide row.
            worksheet4.set_row(row, nil, nil, 1)
        end

        worksheet4.write(row, 0, row_data)
        row += 1
    end


    ###############################################################################
    #
    #
    # Example 5. Autofilter with filter for blanks.
    #

    # Create a blank cell in our test data.
    data[5][0] = ''

    worksheet5.autofilter('A1:D51')
    worksheet5.filter_column('A', 'x == Blanks')

    #
    # Hide the rows that don't match the filter criteria.
    #
    row = 1

    data.each do |row_data|
        region = row_data[0]

        if region == ''
            # Row is visible.
        else
            # Hide row.
            worksheet5.set_row(row, nil, nil, 1)
        end

        worksheet5.write(row, 0, row_data)
        row += 1
    end


    ###############################################################################
    #
    #
    # Example 6. Autofilter with filter for non-blanks.
    #

    worksheet6.autofilter('A1:D51')
    worksheet6.filter_column('A', 'x == NonBlanks')

    #
    # Hide the rows that don't match the filter criteria.
    #
    row = 1

    data.each do |row_data|
        region = row_data[0]

        if region != ''
            # Row is visible.
        else
            # Hide row.
            worksheet6.set_row(row, nil, nil, 1)
        end

        worksheet6.write(row, 0, row_data)
        row += 1
    end

    workbook.close

    # do assertion
    compare_file("#{PERL_OUTDIR}/autofilter.xls", @file)
  end

  def get_data_for_autofilter
    [
      ['East',      'Apple',     9000,      'July'],
      ['East',      'Apple',     5000,      'July'],
      ['South',     'Orange',    9000,      'September'],
      ['North',     'Apple',     2000,      'November'],
      ['West',      'Apple',     9000,      'November'],
      ['South',     'Pear',      7000,      'October'],
      ['North',     'Pear',      9000,      'August'],
      ['West',      'Orange',    1000,      'December'],
      ['West',      'Grape',     1000,      'November'],
      ['South',     'Pear',      10000,     'April'],
      ['West',      'Grape',     6000,      'January'],
      ['South',     'Orange',    3000,      'May'],
      ['North',     'Apple',     3000,      'December'],
      ['South',     'Apple',     7000,      'February'],
      ['West',      'Grape',     1000,      'December'],
      ['East',      'Grape',     8000,      'February'],
      ['South',     'Grape',     10000,     'June'],
      ['West',      'Pear',      7000,      'December'],
      ['South',     'Apple',     2000,      'October'],
      ['East',      'Grape',     7000,      'December'],
      ['North',     'Grape',     6000,      'April'],
      ['East',      'Pear',      8000,      'February'],
      ['North',     'Apple',     7000,      'August'],
      ['North',     'Orange',    7000,      'July'],
      ['North',     'Apple',     6000,      'June'],
      ['South',     'Grape',     8000,      'September'],
      ['West',      'Apple',     3000,      'October'],
      ['South',     'Orange',    10000,     'November'],
      ['West',      'Grape',     4000,      'July'],
      ['North',     'Orange',    5000,      'August'],
      ['East',      'Orange',    1000,      'November'],
      ['East',      'Orange',    4000,      'October'],
      ['North',     'Grape',     5000,      'August'],
      ['East',      'Apple',     1000,      'December'],
      ['South',     'Apple',     10000,     'March'],
      ['East',      'Grape',     7000,      'October'],
      ['West',      'Grape',     1000,      'September'],
      ['East',      'Grape',     10000,     'October'],
      ['South',     'Orange',    8000,      'March'],
      ['North',     'Apple',     4000,      'July'],
      ['South',     'Orange',    5000,      'July'],
      ['West',      'Apple',     4000,      'June'],
      ['East',      'Apple',     5000,      'April'],
      ['North',     'Pear',      3000,      'August'],
      ['East',      'Grape',     9000,      'November'],
      ['North',     'Orange',    8000,      'October'],
      ['East',      'Apple',     10000,     'June'],
      ['South',     'Pear',      1000,      'December'],
      ['North',     'Grape',     10000,     'July'],
      ['East',      'Grape',     6000,      'February'],
    ]
  end

  def test_regions
    workbook = WriteExcel.new(@file)

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
    compare_file("#{PERL_OUTDIR}/regions.xls", @file)
  end

  def test_stats
    workbook = WriteExcel.new(@file)
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
    compare_file("#{PERL_OUTDIR}/stats.xls", @file)
  end

  def test_hyperlink1
    # Create a new workbook and add a worksheet
    workbook  = WriteExcel.new(@file)
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
    compare_file("#{PERL_OUTDIR}/hyperlink.xls", @file)
  end

  def test_copyformat
    # Create workbook1
    workbook1       = WriteExcel.new(@file)
    worksheet1      = workbook1.add_worksheet
    format1a        = workbook1.add_format
    format1b        = workbook1.add_format

    # Create workbook2
    file2 = StringIO.new
    workbook2       = WriteExcel.new(file2)
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
    compare_file("#{PERL_OUTDIR}/workbook1.xls", @file)
    compare_file("#{PERL_OUTDIR}/workbook2.xls", file2)
  end

  def test_data_validate
    workbook  = WriteExcel.new(@file)
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
    heading1 = 'Some examples of data validation in WriteExcel'
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
    compare_file("#{PERL_OUTDIR}/data_validate.xls", @file)
  end

  def test_merge1
    workbook  = WriteExcel.new(@file)
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
    compare_file("#{PERL_OUTDIR}/merge1.xls", @file)
  end

  def test_merge2
    workbook  = WriteExcel.new(@file)
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
    compare_file("#{PERL_OUTDIR}/merge2.xls", @file)
  end

  def test_merge3
    workbook  = WriteExcel.new(@file)
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
    compare_file("#{PERL_OUTDIR}/merge3.xls", @file)
  end

  def test_merge4
    # Create a new workbook and add a worksheet
    workbook  = WriteExcel.new(@file)
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
    compare_file("#{PERL_OUTDIR}/merge4.xls", @file)
  end

  def test_merge5
    # Create a new workbook and add a worksheet
    workbook  = WriteExcel.new(@file)
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


    worksheet.merge_range('D4:D9', 'Rotation 90', format2)



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


    worksheet.merge_range('F4:F9', 'Rotation -90', format3)

    workbook.close

    # do assertion
    compare_file("#{PERL_OUTDIR}/merge5.xls", @file)
  end

  def test_images
    # Create a new workbook called simple.xls and add a worksheet
    workbook   = WriteExcel.new(@file)
    worksheet1 = workbook.add_worksheet('Image 1')
    worksheet2 = workbook.add_worksheet('Image 2')
    worksheet3 = workbook.add_worksheet('Image 3')
    worksheet4 = workbook.add_worksheet('Image 4')
    bp=1

    # Insert a basic image
    worksheet1.write('A10', "Image inserted into worksheet.")
    worksheet1.insert_image('A1', File.join(TEST_DIR,'republic.png'))


    # Insert an image with an offset
    worksheet2.write('A10', "Image inserted with an offset.")
    worksheet2.insert_image('A1', File.join(TEST_DIR,'republic.png'), 32, 10)

    # Insert a scaled image
    worksheet3.write('A10', "Image scaled: width x 2, height x 0.8.")
    worksheet3.insert_image('A1', File.join(TEST_DIR,'republic.png'), 0, 0, 2, 0.8)

    # Insert an image over varied column and row sizes
    # This does not require any additional work

    # Set the cols and row sizes
    # NOTE: you must do this before you call insert_image()
    worksheet4.set_column('A:A', 5)
    worksheet4.set_column('B:B', nil, nil, 1) # Hidden
    worksheet4.set_column('C:D', 10)
    worksheet4.set_row(0, 30)
    worksheet4.set_row(3, 5)

    worksheet4.write('A10', "Image inserted over scaled rows and columns.")
    worksheet4.insert_image('A1', File.join(TEST_DIR,'republic.png'))

    workbook.close

    # do assertion
    compare_file("#{PERL_OUTDIR}/images.xls", @file)
  end

  def test_tab_colors
    workbook   = WriteExcel.new(@file)

    worksheet1 =  workbook.add_worksheet
    worksheet2 =  workbook.add_worksheet
    worksheet3 =  workbook.add_worksheet
    worksheet4 =  workbook.add_worksheet

    # Worsheet1 will have the default tab colour.
    worksheet2.set_tab_color('red')
    worksheet3.set_tab_color('green')
    worksheet4.set_tab_color(0x35) # Orange

    workbook.close

    # do assertion
    compare_file("#{PERL_OUTDIR}/tab_colors.xls", @file)
  end

  def test_stocks
    # Create a new workbook and add a worksheet
    workbook  = WriteExcel.new(@file)
    worksheet = workbook.add_worksheet

    # Set the column width for columns 1, 2, 3 and 4
    worksheet.set_column(0, 3, 15)


    # Create a format for the column headings
    header = workbook.add_format
    header.set_bold
    header.set_size(12)
    header.set_color('blue')


    # Create a format for the stock price
    f_price = workbook.add_format
    f_price.set_align('left')
    f_price.set_num_format('$0.00')


    # Create a format for the stock volume
    f_volume = workbook.add_format
    f_volume.set_align('left')
    f_volume.set_num_format('#,##0')


    # Create a format for the price change. This is an example of a conditional
    # format. The number is formatted as a percentage. If it is positive it is
    # formatted in green, if it is negative it is formatted in red and if it is
    # zero it is formatted as the default font colour (in this case black).
    # Note: the [Green] format produces an unappealing lime green. Try
    # [Color 10] instead for a dark green.
    #
    f_change = workbook.add_format
    f_change.set_align('left')
    f_change.set_num_format('[Green]0.0%;[Red]-0.0%;0.0%')


    # Write out the data
    worksheet.write(0, 0, 'Company', header)
    worksheet.write(0, 1, 'Price',   header)
    worksheet.write(0, 2, 'Volume',  header)
    worksheet.write(0, 3, 'Change',  header)

    worksheet.write(1, 0, 'Damage Inc.'     )
    worksheet.write(1, 1, 30.25,     f_price)  # $30.25
    worksheet.write(1, 2, 1234567,   f_volume) # 1,234,567
    worksheet.write(1, 3, 0.085,     f_change) # 8.5% in green

    worksheet.write(2, 0, 'Dump Corp.'      )
    worksheet.write(2, 1, 1.56,      f_price)  # $1.56
    worksheet.write(2, 2, 7564,      f_volume) # 7,564
    worksheet.write(2, 3, -0.015,    f_change) # -1.5% in red

    worksheet.write(3, 0, 'Rev Ltd.'        )
    worksheet.write(3, 1, 0.13,      f_price)  # $0.13
    worksheet.write(3, 2, 321,       f_volume) # 321
    worksheet.write(3, 3, 0,         f_change) # 0 in the font color (black)

    workbook.close

    # do assertion
    compare_file("#{PERL_OUTDIR}/stocks.xls", @file)
  end

  def test_protection
    workbook  = WriteExcel.new(@file)
    worksheet = workbook.add_worksheet

    # Create some format objects
    locked    = workbook.add_format(:locked => 1)
    unlocked  = workbook.add_format(:locked => 0)
    hidden    = workbook.add_format(:hidden => 1)

    # Format the columns
    worksheet.set_column('A:A', 42)
    worksheet.set_selection('B3:B3')

    # Protect the worksheet
    worksheet.protect

    # Examples of cell locking and hiding
    worksheet.write('A1', 'Cell B1 is locked. It cannot be edited.')
    worksheet.write('B1', '=1+2', locked)

    worksheet.write('A2', 'Cell B2 is unlocked. It can be edited.')
    worksheet.write('B2', '=1+2', unlocked)

    worksheet.write('A3', "Cell B3 is hidden. The formula isn't visible.")
    worksheet.write('B3', '=1+2', hidden)

    worksheet.write('A5', 'Use Menu->Tools->Protection->Unprotect Sheet')
    worksheet.write('A6', 'to remove the worksheet protection.   ')

    workbook.close

    # do assertion
    compare_file("#{PERL_OUTDIR}/protection.xls", @file)
  end

  def test_date_time
    # Create a new workbook and add a worksheet
    workbook  = WriteExcel.new(@file)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    # Expand the first column so that the date is visible.
    worksheet.set_column("A:B", 30)

    # Write the column headers
    worksheet.write('A1', 'Formatted date', bold)
    worksheet.write('B1', 'Format',         bold)

    # Examples date and time formats. In the output file compare how changing
    # the format codes change the appearance of the date.
    #
    date_formats = [
        'dd/mm/yy',
        'mm/dd/yy',
        '',
        'd mm yy',
        'dd mm yy',
        '',
        'dd m yy',
        'dd mm yy',
        'dd mmm yy',
        'dd mmmm yy',
        '',
        'dd mm y',
        'dd mm yyy',
        'dd mm yyyy',
        '',
        'd mmmm yyyy',
        '',
        'dd/mm/yy',
        'dd/mm/yy hh:mm',
        'dd/mm/yy hh:mm:ss',
        'dd/mm/yy hh:mm:ss.000',
        '',
        'hh:mm',
        'hh:mm:ss',
        'hh:mm:ss.000',
    ]

    # Write the same date and time using each of the above formats. The empty
    # string formats create a blank line to make the example clearer.
    #
    row = 0
    date_formats.each do |date_format|
      row += 1
      next if date_format == ''

      # Create a format for the date or time.
      format =  workbook.add_format(
                                  :num_format => date_format,
                                  :align      => 'left'
                                 )

      # Write the same date using different formats.
      worksheet.write_date_time(row, 0, '2004-08-01T12:30:45.123', format)
      worksheet.write(row, 1, date_format)
    end

    # The following is an example of an invalid date. It is written as a string instead
    # of a number. This is also Excel's default behaviour.
    #
    row += 2
    worksheet.write_date_time(row, 0, '2004-13-01T12:30:45.123')
    worksheet.write(row, 1, 'Invalid date. Written as string.', bold)

    workbook.close

    # do assertion
    compare_file("#{PERL_OUTDIR}/date_time.xls", @file)
  end

  def test_diag_border
    workbook  = WriteExcel.new(@file)
    worksheet = workbook.add_worksheet

    format1   = workbook.add_format(:diag_type     => 1)
    format2   = workbook.add_format(:diag_type     => 2)
    format3   = workbook.add_format(:diag_type     => 3)
    format4   = workbook.add_format(
                                  :diag_type       => 3,
                                  :diag_border     => 7,
                                  :diag_color      => 'red'
                )

    worksheet.write('B3',  'Text', format1)
    worksheet.write('B6',  'Text', format2)
    worksheet.write('B9',  'Text', format3)
    worksheet.write('B12', 'Text', format4)

    workbook.close

    # do assertion
    compare_file("#{PERL_OUTDIR}/diag_border.xls", @file)
  end

  def test_headers
    workbook  = WriteExcel.new(@file)
    preview   = "Select Print Preview to see the header and footer"


    ######################################################################
    #
    # A simple example to start
    #
    worksheet1  = workbook.add_worksheet('Simple')

    header1     = '&CHere is some centred text.'

    footer1     = '&LHere is some left aligned text.'


    worksheet1.set_header(header1)
    worksheet1.set_footer(footer1)

    worksheet1.set_column('A:A', 50)
    worksheet1.write('A1', preview)


    ######################################################################
    #
    # This is an example of some of the header/footer variables.
    #
    worksheet2  = workbook.add_worksheet('Variables')

    header2     = '&LPage &P of &N'+
                      '&CFilename: &F' +
                      '&RSheetname: &A'

    footer2     = '&LCurrent date: &D'+
                      '&RCurrent time: &T'

    worksheet2.set_header(header2)
    worksheet2.set_footer(footer2)


    worksheet2.set_column('A:A', 50)
    worksheet2.write('A1', preview)
    worksheet2.write('A21', "Next sheet")
    worksheet2.set_h_pagebreaks(20)


    ######################################################################
    #
    # This example shows how to use more than one font
    #
    worksheet3 = workbook.add_worksheet('Mixed fonts')

    header3    = '&C' +
                     '&"Courier New,Bold"Hello ' +
                     '&"Arial,Italic"World'

    footer3    = '&C' +
                     '&"Symbol"e' +
                     '&"Arial" = mc&X2'

    worksheet3.set_header(header3)
    worksheet3.set_footer(footer3)

    worksheet3.set_column('A:A', 50)
    worksheet3.write('A1', preview)


    ######################################################################
    #
    # Example of line wrapping
    #
    worksheet4 = workbook.add_worksheet('Word wrap')

    header4    = "&CHeading 1\nHeading 2\nHeading 3"

    worksheet4.set_header(header4)

    worksheet4.set_column('A:A', 50)
    worksheet4.write('A1', preview)


    ######################################################################
    #
    # Example of inserting a literal ampersand &
    #
    worksheet5 = workbook.add_worksheet('Ampersand')

    header5    = "&CCuriouser && Curiouser - Attorneys at Law"

    worksheet5.set_header(header5)

    worksheet5.set_column('A:A', 50)
    worksheet5.write('A1', preview)

    workbook.close

    # do assertion
    compare_file("#{PERL_OUTDIR}/headers.xls", @file)
  end

  def test_demo
    workbook   = WriteExcel.new(@file)
worksheet  = workbook.add_worksheet('Demo')
worksheet2 = workbook.add_worksheet('Another sheet')
worksheet3 = workbook.add_worksheet('And another')

bold       = workbook.add_format(:bold => 1)

#######################################################################
#
# Write a general heading
#
worksheet.set_column('A:A', 36, bold)
worksheet.set_column('B:B', 20       )
worksheet.set_row(0,     40       )

heading  = workbook.add_format(
                                :bold    => 1,
                                :color   => 'blue',
                                :size    => 16,
                                :merge   => 1,
                                :align  => 'vcenter'
                              )

headings = ['Features of Spreadsheet::WriteExcel', '']
worksheet.write_row('A1', headings, heading)


#######################################################################
#
# Some text examples
#
text_format  = workbook.add_format(
                                    :bold    => 1,
                                    :italic  => 1,
                                    :color   => 'red',
                                    :size    => 18,
                                    :font    =>'Lucida Calligraphy'
                                  )

# A phrase in Cyrillic
unicode = [
            "042d0442043e002004440440043004370430002004"+
            "3d043000200440044304410441043a043e043c0021"
          ].pack('H*')

worksheet.write('A2', "Text")
worksheet.write('B2', "Hello Excel")
worksheet.write('A3', "Formatted text")
worksheet.write('B3', "Hello Excel", text_format)
worksheet.write('A4', "Unicode text")
worksheet.write_utf16be_string('B4', unicode)


#######################################################################
#
# Some numeric examples
#
num1_format  = workbook.add_format(:num_format => '$#,##0.00')
num2_format  = workbook.add_format(:num_format => ' d mmmm yyy')

worksheet.write('A5', "Numbers")
worksheet.write('B5', 1234.56)
worksheet.write('A6', "Formatted numbers")
worksheet.write('B6', 1234.56, num1_format)
worksheet.write('A7', "Formatted numbers")
worksheet.write('B7', 37257, num2_format)


#######################################################################
#
# Formulae
#
worksheet.set_selection('B8')
worksheet.write('A8', 'Formulas and functions, "=SIN(PI()/4)"')
worksheet.write('B8', '=SIN(PI()/4)')


#######################################################################
#
# Hyperlinks
#
worksheet.write('A9', "Hyperlinks")
worksheet.write('B9',  'http://www.perl.com/' )


#######################################################################
#
# Images
#
worksheet.write('A10', "Images")
worksheet.insert_image('B10', "#{TEST_DIR}/republic.png", 16, 8)


#######################################################################
#
# Misc
#
worksheet.write('A18', "Page/printer setup")
worksheet.write('A19', "Multiple worksheets")

workbook.close

    # do assertion
    compare_file("#{PERL_OUTDIR}/demo.xls", @file)
  end

  def test_unicode_cyrillic
    # Create a Russian worksheet name in utf8.
    sheet   = [0x0421, 0x0442, 0x0440, 0x0430, 0x043D, 0x0438,
                         0x0446, 0x0430].pack("U*")

    # Create a Russian string.
    str     = [0x0417, 0x0434, 0x0440, 0x0430, 0x0432, 0x0441,
                       0x0442, 0x0432, 0x0443, 0x0439, 0x0020, 0x041C,
                       0x0438, 0x0440, 0x0021].pack("U*")

    workbook  = WriteExcel.new(@file)
    worksheet = workbook.add_worksheet(sheet + '1')

    worksheet.set_column('A:A', 18)
    worksheet.write('A1', str)

    workbook.close

    # do assertion
    compare_file("#{PERL_OUTDIR}/unicode_cyrillic.xls", @file)
  end

  def test_defined_name
    workbook   = WriteExcel.new(@file)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet

    workbook.define_name('Exchange_rate', '=0.96')
    workbook.define_name('Sales',         '=Sheet1!$G$1:$H$10')
    workbook.define_name('Sheet2!Sales',  '=Sheet2!$G$1:$G$10')

    workbook.sheets.each do |worksheet|
      worksheet.set_column('A:A', 45)
      worksheet.write('A2', 'This worksheet contains some defined names,')
      worksheet.write('A3', 'See the Insert -> Name -> Define dialog.')
    end

    worksheet1.write('A4', '=Exchange_rate')

    workbook.close

    # do assertion
    compare_file("#{PERL_OUTDIR}/defined_name.xls", @file)
  end

  def test_chart_area
workbook  = WriteExcel.new(@file)
worksheet = workbook.add_worksheet
bold      = workbook.add_format(:bold => 1)

# Add the data to the worksheet that the charts will refer to.
headings = [ 'Category', 'Values 1', 'Values 2' ]
data = [
    [ 2, 3, 4, 5, 6, 7 ],
    [ 1, 4, 5, 2, 1, 5 ],
    [ 3, 6, 7, 5, 4, 3 ]
]

worksheet.write('A1', headings, bold)
worksheet.write('A2', data)


###############################################################################
#
# Example 1. A minimal chart.
#
chart1 = workbook.add_chart(:type => Chart::Area)

# Add values only. Use the default categories.
chart1.add_series( :values => '=Sheet1!$B$2:$B$7' )

###############################################################################
#
# Example 2. A minimal chart with user specified categories (X axis)
#            and a series name.
#
chart2 = workbook.add_chart(:type => Chart::Area)

# Configure the series.
chart2.add_series(
    :categories => '=Sheet1!$A$2:$A$7',
    :values     => '=Sheet1!$B$2:$B$7',
    :name       => 'Test data series 1'
)

###############################################################################
#
# Example 3. Same as previous chart but with added title and axes labels.
#
chart3 = workbook.add_chart(:type => Chart::Area)

# Configure the series.
chart3.add_series(
    :categories => '=Sheet1!$A$2:$A$7',
    :values     => '=Sheet1!$B$2:$B$7',
    :name       => 'Test data series 1'
)

# Add some labels.
chart3.set_title( :name => 'Results of sample analysis' )
chart3.set_x_axis( :name => 'Sample number' )
chart3.set_y_axis( :name => 'Sample length (cm)' )

###############################################################################
#
# Example 4. Same as previous chart but with an added series
#
chart4 = workbook.add_chart(:name => 'Results Chart', :type => Chart::Area)

# Configure the series.
chart4.add_series(
    :categories => '=Sheet1!$A$2:$A$7',
    :values     => '=Sheet1!$B$2:$B$7',
    :name       => 'Test data series 1'
)

# Add another series.
chart4.add_series(
    :categories => '=Sheet1!$A$2:$A$7',
    :values     => '=Sheet1!$C$2:$C$7',
    :name       => 'Test data series 2'
)

# Add some labels.
chart4.set_title( :name => 'Results of sample analysis' )
chart4.set_x_axis( :name => 'Sample number' )
chart4.set_y_axis( :name => 'Sample length (cm)' )

###############################################################################
#
# Example 5. Same as Example 3 but as an embedded chart.
#
chart5 = workbook.add_chart(:type => Chart::Area, :embedded => 1)

# Configure the series.
chart5.add_series(
  :categories => '=Sheet1!$A$2:$A$7',
  :values     => '=Sheet1!$B$2:$B$7',
  :name       => 'Test data series 1'
)

# Add some labels.
chart5.set_title(:name => 'Results of sample analysis' )
chart5.set_x_axis(:name => 'Sample number')
chart5.set_y_axis(:name => 'Sample length (cm)')

# Insert the chart into the main worksheet.
worksheet.insert_chart('E2', chart5)

# File save
workbook.close

    # do assertion
    compare_file("#{PERL_OUTDIR}/chart_area.xls", @file)
  end

  def test_chart_bar
workbook  = WriteExcel.new(@file)
worksheet = workbook.add_worksheet
bold      = workbook.add_format(:bold => 1)

# Add the data to the worksheet that the charts will refer to.
headings = [ 'Category', 'Values 1', 'Values 2' ]
data = [
    [ 2, 3, 4, 5, 6, 7 ],
    [ 1, 4, 5, 2, 1, 5 ],
    [ 3, 6, 7, 5, 4, 3 ]
]

worksheet.write('A1', headings, bold)
worksheet.write('A2', data)


###############################################################################
#
# Example 1. A minimal chart.
#
chart1 = workbook.add_chart(:type => Chart::Bar)

# Add values only. Use the default categories.
chart1.add_series( :values => '=Sheet1!$B$2:$B$7' )

###############################################################################
#
# Example 2. A minimal chart with user specified categories (X axis)
#            and a series name.
#
chart2 = workbook.add_chart(:type => Chart::Bar)

# Configure the series.
chart2.add_series(
    :categories => '=Sheet1!$A$2:$A$7',
    :values     => '=Sheet1!$B$2:$B$7',
    :name       => 'Test data series 1'
)

###############################################################################
#
# Example 3. Same as previous chart but with added title and axes labels.
#
chart3 = workbook.add_chart(:type => Chart::Bar)

# Configure the series.
chart3.add_series(
    :categories => '=Sheet1!$A$2:$A$7',
    :values     => '=Sheet1!$B$2:$B$7',
    :name       => 'Test data series 1'
)

# Add some labels.
chart3.set_title( :name => 'Results of sample analysis' )
chart3.set_x_axis( :name => 'Sample number' )
chart3.set_y_axis( :name => 'Sample length (cm)' )

###############################################################################
#
# Example 4. Same as previous chart but with an added series
#
chart4 = workbook.add_chart(:name => 'Results Chart', :type => Chart::Bar)

# Configure the series.
chart4.add_series(
    :categories => '=Sheet1!$A$2:$A$7',
    :values     => '=Sheet1!$B$2:$B$7',
    :name       => 'Test data series 1'
)

# Add another series.
chart4.add_series(
    :categories => '=Sheet1!$A$2:$A$7',
    :values     => '=Sheet1!$C$2:$C$7',
    :name       => 'Test data series 2'
)

# Add some labels.
chart4.set_title( :name => 'Results of sample analysis' )
chart4.set_x_axis( :name => 'Sample number' )
chart4.set_y_axis( :name => 'Sample length (cm)' )

###############################################################################
#
# Example 5. Same as Example 3 but as an embedded chart.
#
chart5 = workbook.add_chart(:type => Chart::Bar, :embedded => 1)

# Configure the series.
chart5.add_series(
  :categories => '=Sheet1!$A$2:$A$7',
  :values     => '=Sheet1!$B$2:$B$7',
  :name       => 'Test data series 1'
)

# Add some labels.
chart5.set_title(:name => 'Results of sample analysis' )
chart5.set_x_axis(:name => 'Sample number')
chart5.set_y_axis(:name => 'Sample length (cm)')

# Insert the chart into the main worksheet.
worksheet.insert_chart('E2', chart5)

# File save
workbook.close

    # do assertion
    compare_file("#{PERL_OUTDIR}/chart_bar.xls", @file)
  end

  def test_chart_column
workbook  = WriteExcel.new(@file)
worksheet = workbook.add_worksheet
bold      = workbook.add_format(:bold => 1)

# Add the data to the worksheet that the charts will refer to.
headings = [ 'Category', 'Values 1', 'Values 2' ]
data = [
    [ 2, 3, 4, 5, 6, 7 ],
    [ 1, 4, 5, 2, 1, 5 ],
    [ 3, 6, 7, 5, 4, 3 ]
]

worksheet.write('A1', headings, bold)
worksheet.write('A2', data)


###############################################################################
#
# Example 1. A minimal chart.
#
chart1 = workbook.add_chart(:type => Chart::Column)

# Add values only. Use the default categories.
chart1.add_series( :values => '=Sheet1!$B$2:$B$7' )

###############################################################################
#
# Example 2. A minimal chart with user specified categories (X axis)
#            and a series name.
#
chart2 = workbook.add_chart(:type => Chart::Column)

# Configure the series.
chart2.add_series(
    :categories => '=Sheet1!$A$2:$A$7',
    :values     => '=Sheet1!$B$2:$B$7',
    :name       => 'Test data series 1'
)

###############################################################################
#
# Example 3. Same as previous chart but with added title and axes labels.
#
chart3 = workbook.add_chart(:type => Chart::Column)

# Configure the series.
chart3.add_series(
    :categories => '=Sheet1!$A$2:$A$7',
    :values     => '=Sheet1!$B$2:$B$7',
    :name       => 'Test data series 1'
)

# Add some labels.
chart3.set_title( :name => 'Results of sample analysis' )
chart3.set_x_axis( :name => 'Sample number' )
chart3.set_y_axis( :name => 'Sample length (cm)' )

###############################################################################
#
# Example 4. Same as previous chart but with an added series
#
chart4 = workbook.add_chart(:name => 'Results Chart', :type => Chart::Column)

# Configure the series.
chart4.add_series(
    :categories => '=Sheet1!$A$2:$A$7',
    :values     => '=Sheet1!$B$2:$B$7',
    :name       => 'Test data series 1'
)

# Add another series.
chart4.add_series(
    :categories => '=Sheet1!$A$2:$A$7',
    :values     => '=Sheet1!$C$2:$C$7',
    :name       => 'Test data series 2'
)

# Add some labels.
chart4.set_title( :name => 'Results of sample analysis' )
chart4.set_x_axis( :name => 'Sample number' )
chart4.set_y_axis( :name => 'Sample length (cm)' )

###############################################################################
#
# Example 5. Same as Example 3 but as an embedded chart.
#
chart5 = workbook.add_chart(:type => Chart::Column, :embedded => 1)

# Configure the series.
chart5.add_series(
  :categories => '=Sheet1!$A$2:$A$7',
  :values     => '=Sheet1!$B$2:$B$7',
  :name       => 'Test data series 1'
)

# Add some labels.
chart5.set_title(:name => 'Results of sample analysis' )
chart5.set_x_axis(:name => 'Sample number')
chart5.set_y_axis(:name => 'Sample length (cm)')

# Insert the chart into the main worksheet.
worksheet.insert_chart('E2', chart5)

# File save
workbook.close

    # do assertion
    compare_file("#{PERL_OUTDIR}/chart_column.xls", @file)
  end

  def test_chart_line
workbook  = WriteExcel.new(@file)
worksheet = workbook.add_worksheet
bold      = workbook.add_format(:bold => 1)

# Add the data to the worksheet that the charts will refer to.
headings = [ 'Category', 'Values 1', 'Values 2' ]
data = [
    [ 2, 3, 4, 5, 6, 7 ],
    [ 1, 4, 5, 2, 1, 5 ],
    [ 3, 6, 7, 5, 4, 3 ]
]

worksheet.write('A1', headings, bold)
worksheet.write('A2', data)


###############################################################################
#
# Example 1. A minimal chart.
#
chart1 = workbook.add_chart(:type => Chart::Line)

# Add values only. Use the default categories.
chart1.add_series( :values => '=Sheet1!$B$2:$B$7' )

###############################################################################
#
# Example 2. A minimal chart with user specified categories (X axis)
#            and a series name.
#
chart2 = workbook.add_chart(:type => Chart::Line)

# Configure the series.
chart2.add_series(
    :categories => '=Sheet1!$A$2:$A$7',
    :values     => '=Sheet1!$B$2:$B$7',
    :name       => 'Test data series 1'
)

###############################################################################
#
# Example 3. Same as previous chart but with added title and axes labels.
#
chart3 = workbook.add_chart(:type => Chart::Line)

# Configure the series.
chart3.add_series(
    :categories => '=Sheet1!$A$2:$A$7',
    :values     => '=Sheet1!$B$2:$B$7',
    :name       => 'Test data series 1'
)

# Add some labels.
chart3.set_title( :name => 'Results of sample analysis' )
chart3.set_x_axis( :name => 'Sample number' )
chart3.set_y_axis( :name => 'Sample length (cm)' )

###############################################################################
#
# Example 4. Same as previous chart but with an added series
#
chart4 = workbook.add_chart(:name => 'Results Chart', :type => Chart::Line)

# Configure the series.
chart4.add_series(
    :categories => '=Sheet1!$A$2:$A$7',
    :values     => '=Sheet1!$B$2:$B$7',
    :name       => 'Test data series 1'
)

# Add another series.
chart4.add_series(
    :categories => '=Sheet1!$A$2:$A$7',
    :values     => '=Sheet1!$C$2:$C$7',
    :name       => 'Test data series 2'
)

# Add some labels.
chart4.set_title( :name => 'Results of sample analysis' )
chart4.set_x_axis( :name => 'Sample number' )
chart4.set_y_axis( :name => 'Sample length (cm)' )

###############################################################################
#
# Example 5. Same as Example 3 but as an embedded chart.
#
chart5 = workbook.add_chart(:type => Chart::Line, :embedded => 1)

# Configure the series.
chart5.add_series(
  :categories => '=Sheet1!$A$2:$A$7',
  :values     => '=Sheet1!$B$2:$B$7',
  :name       => 'Test data series 1'
)

# Add some labels.
chart5.set_title(:name => 'Results of sample analysis' )
chart5.set_x_axis(:name => 'Sample number')
chart5.set_y_axis(:name => 'Sample length (cm)')

# Insert the chart into the main worksheet.
worksheet.insert_chart('E2', chart5)

# File save
workbook.close

    # do assertion
    compare_file("#{PERL_OUTDIR}/chart_line.xls", @file)
  end

  def compare_file(expected, target)
    # target is StringIO object.
    assert_equal(
      open(expected, 'rb') { |f| f.read },
      target.string
    )
  end
end
