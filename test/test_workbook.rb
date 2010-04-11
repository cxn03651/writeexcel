# -*- coding: utf-8 -*-
require 'helper'
require "stringio"

class TC_Workbook < Test::Unit::TestCase

  def setup
    @test_file  = StringIO.new
    @workbook   = Workbook.new(@test_file)
  end

  def teardown
    @workbook.close
  end

  def test_new
    assert_kind_of(Workbook, @workbook)
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

  def test_set_tempdir_after_sheet_added
    # after shees added, call set_tempdir raise RuntimeError
    @workbook.add_worksheet('name')
    assert_raise(RuntimeError, "already sheet exists, but set_tempdir() doesn't raise"){
      @workbook.set_tempdir
    }
  end

  def test_set_tempdir_with_invalid_dir
    # invalid dir raise RuntimeError
    while true do
      dir = Time.now.to_s
      break unless FileTest.directory?(dir)
      sleep 0.1
    end
    assert_raise(RuntimeError, "set_tempdir() doesn't raise invalid dir:#{dir}."){
      @workbook.set_tempdir(dir)
    }
  end

=begin
#
# Comment out because Workbook#check_sheetname was set to private method.
#
  def test_check_sheetname
    valids   = valid_sheetname
    invalids = invalid_sheetname
    worksheet1 = @workbook.add_worksheet              # implicit name 'Sheet1'
    worksheet2 = @workbook.add_worksheet              # implicit name 'Sheet2'
    worksheet3 = @workbook.add_worksheet 'Sheet3'     # implicit name 'Sheet3'
    worksheet1 = @workbook.add_worksheet 'Sheetz'     # implicit name 'Sheetz'

    valids.each do |test|
      target    = test[0]
      sheetname = test[1]
      caption   = test[2]
      assert_nothing_raised { @workbook.check_sheetname(sheetname) }
    end
    invalids.each do |test|
      target    = test[0]
      sheetname = test[1]
      caption   = test[2]
      assert_raise(RuntimeError, "sheetname: #{sheetname}") {
          @workbook.check_sheetname(sheetname)
        }
    end
  end
=end

  def test_raise_set_compatibility_after_sheet_creation
    @workbook.add_worksheet
    assert_raise(RuntimeError) { @workbook.compatibility_mode }
  end

  def valid_sheetname
    [
      # Tests for valid names
      [ 'PASS', nil,        'No worksheet name'           ],
      [ 'PASS', '',         'Blank worksheet name'        ],
      [ 'PASS', 'Sheet10',  'Valid worksheet name'        ],
      [ 'PASS', 'a' * 31,   'Valid 31 char name'          ]
    ]
  end

  def invalid_sheetname
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

end
