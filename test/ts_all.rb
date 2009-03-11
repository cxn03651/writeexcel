$LOAD_PATH.unshift(Dir.pwd)
$LOAD_PATH.unshift(Dir.pwd + "/test")

require "tc_biff"
require "tc_ole"
require "tc_workbook"
require "tc_worksheet"
require "tc_format"
require "tc_formula"
require 'tc_chart'
require "tc_excel"
require "test_00_IEEE_double"
require 'test_01_add_worksheet'
require 'test_02_merge_formats'
require 'test_04_dimensions'
require 'test_05_rows'
require 'test_06_extsst'
require 'test_11_date_time'
require 'test_12_date_only'
require 'test_13_date_seconds.rb'
require 'test_22_mso_drawing_group'
require 'test_23_note'
require 'test_24_txo'
require 'test_26_autofilter'
require 'test_27_autofilter'
require 'test_28_autofilter'
require 'test_29_process_jpg'
require 'test_30_validation_dval'
require 'test_31_validation_dv_strings'
require 'test_32_validation_dv_formula'
