require 'writeexcel'
include Spreadsheet

bp=1
  workbook  = WriteExcel.new("japanese_utf8.xls", :font => 'ＭＳＰ ゴシック', :size => 11)
  worksheet = workbook.add_worksheet()
  worksheet.set_column('B:B', 11)

  fmt_name_label = workbook.add_format
  fmt_name_label.set_align('right')
  fmt_name_label.set_top(2)
  fmt_name_label.set_left(2)
  fmt_name_label.set_right(1)
  fmt_name_label.set_bottom(1)
  worksheet.write('B2', '氏名：', fmt_name_label)

  fmt_center_across = workbook.add_format(:center_across => 1)
  worksheet.write(1, 2, "ここに氏名を入力", fmt_center_across)
  3.upto(8) { |col| worksheet.write_blank(1, col, fmt_center_across) }

  fmt_birthday_label = workbook.add_format
  fmt_birthday_label.set_align('right')
  fmt_birthday_label.set_top(1)
  fmt_birthday_label.set_left(2)
  fmt_birthday_label.set_right(1)
  fmt_birthday_label.set_bottom(2)
  worksheet.write('B3', '生年月日：', fmt_birthday_label)

  fmt_ad_label = workbook.add_format
  fmt_ad_label.set_align('right')
  fmt_ad_label.set_top(1)
  fmt_ad_label.set_bottom(2)
  worksheet.write('C3', '西暦', fmt_ad_label)

  worksheet.data_validation('D3',
      :validate      => 'integer',
      :criteria      => 'between',
      :minimum       => 1900,
      :maximum       => 2009,
      :input_title   => '生まれた年',
      :input_message => '西暦４桁で入力してください。'
    )

  worksheet.set_column('E:E', 2.13)
  worksheet.write('E3', '年')

  worksheet.data_validation('F3',
      :validate      => 'list',
      :value         => [1,2,3,4,5,6,7,8,9,10,11,12],
      :input_message => 'リストから選択してください。'
    )
  worksheet.set_column('F:F', 3.75)

  worksheet.set_column('G:G', 2.13)
  worksheet.write('G3', '月')

  worksheet.data_validation('H3',
      :validate      => 'list',
      :value         => Array.new(31) {|i| i + 1},
      :input_message => 'リストから選択してください。'
    )
  worksheet.set_column('H:H', 3.75)
  worksheet.set_column('I:I', 2.13)
  worksheet.write('I3', '日')

  workbook.close
