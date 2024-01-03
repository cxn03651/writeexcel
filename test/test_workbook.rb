# -*- coding: utf-8 -*-
require 'helper'
require "stringio"

class TC_Workbook < Minitest::Test

  def setup
    @test_file  = StringIO.new
    @workbook   = Workbook.new(@test_file)
  end

  def test_new
    assert_kind_of(Workbook, @workbook)
  end

  def test_new_with_block
    test_file = StringIO.new
    workbook = Workbook.new(test_file) do |book|
      book.add_worksheet('test_block')
    end

    assert_kind_of(Workbook, workbook)
    assert_true(test_file.closed?)
  end

  def test_add_worksheet
    sheetnames = ['sheet1', 'sheet2']
    (0 .. sheetnames.size-1).each do |i|
      sheets = @workbook.sheets
      assert_equal(i, sheets.size)
      @workbook.add_worksheet(sheetnames[i])
      sheets = @workbook.sheets
      assert_equal(i+1, sheets.size)
    end
  end

  def valid_sheetnames
    [
      # Tests for valid names
      [ 'PASS', nil,        'No worksheet name'           ],
      [ 'PASS', '',         'Blank worksheet name'        ],
      [ 'PASS', 'Sheet10',  'Valid worksheet name'        ],
      [ 'PASS', 'a' * 31,   'Valid 31 char name'          ]
    ]
  end

  def invalid_sheetnames
    [
      # Tests for invalid names
      [ 'FAIL', 'Sheet1',   'Caught duplicate name'       ],
      [ 'FAIL', 'Sheet2',   'Caught duplicate name'       ],
      [ 'FAIL', 'Sheet3',   'Caught duplicate name'       ],
      [ 'FAIL', 'sheet1',   'Caught case-insensitive name'],
      [ 'FAIL', 'SHEET1',   'Caught case-insensitive name'],
      [ 'FAIL', 'sheetz',   'Caught case-insensitive name'],
      [ 'FAIL', 'SHEETZ',   'Caught case-insensitive name'],
      [ 'FAIL', 'a' * 32,   'Caught long name'            ],
      [ 'FAIL', '[',        'Caught invalid char'         ],
      [ 'FAIL', ']',        'Caught invalid char'         ],
      [ 'FAIL', ':',        'Caught invalid char'         ],
      [ 'FAIL', '*',        'Caught invalid char'         ],
      [ 'FAIL', '?',        'Caught invalid char'         ],
      [ 'FAIL', '/',        'Caught invalid char'         ],
      [ 'FAIL', '\\',       'Caught invalid char'         ]
    ]
  end

  def test_add_format_must_accept_one_or_more_hash_params
    font    = {
      :font   => 'ＭＳ 明朝',
      :size   => 12,
      :color  => 'blue',
      :bold   => 1
    }
    shading = {
      :bg_color => 'green',
      :pattern  => 1
    }
    properties = font.merge(shading)

    format1 = @workbook.add_format(properties)
    format2 = @workbook.add_format(font, shading)
    assert(format_equal?(format1, format2))
  end

  def format_equal?(f1, f2)
    require 'yaml'
    re = /xf_index: \d+\n/
    YAML.dump(f1).sub(re, '') == YAML.dump(f2).sub(re, '')
  end
end
