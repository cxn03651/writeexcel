$LOAD_PATH.unshift(File.dirname(__FILE__))

require "biffwriter"
require "olewriter"
require "workbook"
require "worksheet"
require "format"
require "formula"

module Spreadsheet
  class WriteExcel < Workbook
    VERSION = "0.1.00"
  end
end
