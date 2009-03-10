require 'excel'
include Spreadsheet

def unpack_record(data)
  data.unpack('C*').map! {|c| sprintf("%02X", c) }.join(' ')
end

bp=1
    test_file           = "temp_test_file.xls"
    workbook            = Excel.new(test_file)
    workbook.compatibility_mode(1)
    tests               = []
    
    # for test case 1
    row  = 1
    col1 = 0
    col2 = 0
    worksheet = workbook.add_worksheet
    worksheet.set_row(row, 15)
    tests.push(
                 [
                    " \tset_row(): row = #{row}, col1 = #{col1}, col2 = #{col2}",
                    {
                      :col_min => 0,
                      :col_max => 0,
                    }
                 ]
              )

    workbook.biff_only  = 1


#  dump data

=begin

# this part is PASS.

  print unpack_record(workbook.data) +"\n"

  print 'workbook.store_bof(0x0005)'+"\n"
 workbook.store_bof(0x0005)
  print unpack_record(workbook.data) +"\n"

  print 'workbook.store_codepage'+"\n"
 workbook.store_codepage
  print unpack_record(workbook.data) +"\n"

  print 'workbook.store_window1'+"\n"
 workbook.store_window1
  print unpack_record(workbook.data) +"\n"

  print 'workbook.store_hideobj'+"\n"
 workbook.store_hideobj
  print unpack_record(workbook.data) +"\n"

  print 'workbook.store_1904'+"\n"
 workbook.store_1904
  print unpack_record(workbook.data) +"\n"

  print 'workbook.store_all_fonts'+"\n"
 workbook.store_all_fonts
  print unpack_record(workbook.data) +"\n"

  print 'workbook.store_all_num_formats'+"\n"
 workbook.store_all_num_formats
  print unpack_record(workbook.data) +"\n"

  print 'workbook.store_all_xfs'+"\n"
 workbook.store_all_xfs
  print unpack_record(workbook.data) +"\n"

  print 'workbook.store_all_styles'+"\n"
 workbook.store_all_styles
  print unpack_record(workbook.data) +"\n"

  print 'workbook.store_palette'+"\n"
  workbook.store_palette
  print unpack_record(workbook.data) +"\n"

  workbook.calc_sheet_offsets

  print 'workbook.boundsheet' + "\n"
  workbook.store_boundsheet(worksheet.name,
      worksheet.offset,
      worksheet.type,
      worksheet.hidden,
      worksheet.encoding)
  print unpack_record(workbook.data) +"\n"

  print "workbook.store_country\n"
  workbook.store_country
  print unpack_record(workbook.data) +"\n"
  
  print "workbook.add_mso_drawing_group\n"
  workbook.add_mso_drawing_group
  print unpack_record(workbook.data) +"\n"

  print "workbook.store_shared_strings\n"
  workbook.store_shared_strings
  print unpack_record(workbook.data) +"\n"

  print "workbook.store_extsst\n"
  workbook.store_extsst
  print unpack_record(workbook.data) +"\n"
  
=end

=begin

 # this is worksheet part 1.   pass.

  print "worksheet.store_dimensions\n"
  worksheet.store_dimensions
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_autofilters\n"
  worksheet.store_autofilters
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_autofilterinfo\n"
  worksheet.store_autofilterinfo
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_filtermode\n"
  worksheet.store_filtermode
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_defcol\n"
  worksheet.store_defcol
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_password\n"
  worksheet.store_password
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_protect\n"
  worksheet.store_protect
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_obj_protect\n"
  worksheet.store_obj_protect
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_setup\n"
  worksheet.store_setup
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_margin_bottom\n"
  worksheet.store_margin_bottom
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_margin_top\n"
  worksheet.store_margin_top
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_margin_right\n"
  worksheet.store_margin_right
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_margin_left\n"
  worksheet.store_margin_left
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_vcenter\n"
  worksheet.store_vcenter
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_hcenter\n"
  worksheet.store_hcenter
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_footer\n"
  worksheet.store_footer
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_header\n"
  worksheet.store_header
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_vbreak\n"
  worksheet.store_vbreak
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_hbreak\n"
  worksheet.store_hbreak
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_wsbool\n"
  worksheet.store_wsbool
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_defrow\n"
  worksheet.store_defrow
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_guts\n"
  worksheet.store_guts
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_gridset\n"
  worksheet.store_gridset
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_print_gridlines\n"
  worksheet.store_print_gridlines
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_print_headers\n"
  worksheet.store_print_headers
  print unpack_record(worksheet.data) +"\n"

=end

=begin
  print "worksheet.store_table\n"
  worksheet.store_table
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_images\n"
  worksheet.store_images
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_charts\n"
  worksheet.store_charts
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_filters\n"
  worksheet.store_filters
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_comments\n"
  worksheet.store_comments
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_window2\n"
  worksheet.store_window2
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_page_view\n"
  worksheet.store_page_view
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_zoom\n"
  worksheet.store_zoom
  print unpack_record(worksheet.data) +"\n"

  print "worksheet.store_panes\n"
  worksheet.store_panes(*@panes) if !@panes.nil? && @panes != 0
  print unpack_record(worksheet.data) +"\n"


#print "\n\nworkbook.get_data\n"
  print unpack_record(workbook.get_data) +"\n"
  print unpack_record(worksheet.data) +"\n"


exit
=end
    workbook.close

    rows = []
  
    xlsfile = open(test_file, "rb")
    while header = xlsfile.read(4)
      record, length = header.unpack('vv')
      data = xlsfile.read(length)
    
      #read the row records only
      next unless record == 0x0208
      col_min, col_max = data.unpack('x2 vv')
      
      rows.push(
        {
          :col_min => col_min,
          :col_max => col_max
        }
      )
    end

