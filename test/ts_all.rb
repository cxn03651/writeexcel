$LOAD_PATH.unshift(Dir.pwd)
$LOAD_PATH.unshift(Dir.pwd + "/test")

require "tc_biff"
require "tc_ole"
require "tc_workbook"
require "tc_worksheet"
require "tc_format"
require "tc_formula"
require "tc_excel"
require "test_00_IEEE_double"
