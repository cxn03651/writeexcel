require 'excel'
include Spreadsheet

def unpack_record(data)
  data.unpack('C*').map! {|c| sprintf("%02X", c) }.join(' ')
end

  tests = [
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
  
bp=1
    workbook  = Workbook.new('test.xls')
    worksheet = workbook.add_worksheet
    tests.each do |test|
      column     = test['column']
      expression = test['expression']
      tokens     = worksheet.extract_filter_tokens(expression)
      tokens     = worksheet.parse_filter_expression(expression, tokens)
  
      result     = worksheet.store_autofilter(column, *tokens)
  
      target     = test['data'].join(" ")
  
      caption    = " \tfilter_column(#{column}, '#{expression}')"
  
      result     = unpack_record(result)
      p result
      p target
    end

