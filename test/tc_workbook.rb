#####################################################
# t_workbook.rb
#
# Test suite for the Workbook class (workbook.rb)
# Requires testunit 0.1.8 or greater to run properly
#####################################################
base = File.basename(Dir.pwd)
if base == "test" || base =~ /spreadsheet/i
  Dir.chdir("..") if base == "test"
  $LOAD_PATH.unshift(Dir.pwd + "/lib/spreadsheet")
  Dir.chdir("test") rescue nil
end

require "test/unit"
require "biffwriter"
require "olewriter"
require "workbook"
require "worksheet"
require "format"
require 'formula'

class TC_Workbook < Test::Unit::TestCase

  def setup
    t = Time.now.strftime("%Y%m%d")
    path = "temp#{t}-#{$$}-#{rand(0x100000000).to_s(36)}"
    @test_file           = File.join(Dir.tmpdir, path)
    @wb = Workbook.new(@test_file)
  end

  def teardown
    @wb.close
    File.unlink(@test_file) if FileTest.exist?(@test_file)
  end

  def test_new
    assert_kind_of(Workbook, @wb)
  end

  def test_add_worksheet
    sheetnames = ['sheet1', 'sheet2']
    (0 .. sheetnames.size-1).each do |i|
      sheets = @wb.sheets
      assert_equal(i, sheets.size)
      @wb.add_worksheet(sheetnames[i])
      sheets = @wb.sheets
      assert_equal(i+1, sheets.size)
    end
  end

  def test_set_tempdir_after_sheet_added
    # after shees added, call set_tempdir raise RuntimeError
    @wb.add_worksheet('name')
    assert_raise(RuntimeError, "already sheet exists, but set_tempdir() doesn't raise"){
      @wb.set_tempdir
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
      @wb.set_tempdir(dir)
    }
  end

  def test_check_sheetname
    valids   = valid_sheetname
    invalids = invalid_sheetname
    worksheet1 = @wb.add_worksheet              # implicit name 'Sheet1'
    worksheet2 = @wb.add_worksheet              # implicit name 'Sheet2'
    worksheet3 = @wb.add_worksheet 'Sheet3'     # implicit name 'Sheet3'
    worksheet1 = @wb.add_worksheet 'Sheetz'     # implicit name 'Sheetz'

    valids.each do |test|
      target    = test[0]
      sheetname = test[1]
      caption   = test[2]
      assert_nothing_raised { @wb.check_sheetname(sheetname) }
    end
    invalids.each do |test|
      target    = test[0]
      sheetname = test[1]
      caption   = test[2]
      assert_raise(RuntimeError, "sheetname: #{sheetname}") { @wb.check_sheetname(sheetname) }
    end
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
