require 'excel'
include Spreadsheet

def unpack_record(data)
  data.unpack('C*').map! {|c| sprintf("%02X", c) }.join(' ')
end

  tests = [
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
  
bp=1
    workbook  = Workbook.new('test.xls')
    worksheet = workbook.add_worksheet
    tests.each do |test|
      expression = test[0]
      expected   = test[1]
      result     = worksheet.extract_filter_tokens(expression)
  
      testname   = expression || 'none'
      p result
      p expected
    end

