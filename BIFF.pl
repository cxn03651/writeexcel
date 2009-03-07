use Spreadsheet::WriteExcel;


my $test_file   = 'temp_test_file.xls';
my $workbook    = Spreadsheet::WriteExcel->new($test_file);

my $worksheet1 = $workbook->add_worksheet();
my $count1     = 1;

   $worksheet1->write_comment($_ -1, 0, 'aaa') for 1 .. $count1;

   $workbook->_calc_mso_sizes();

   $target     = join " ",  qw(
                            EB 00 5A 00 0F 00 00 F0 52 00 00 00 00 00 06 F0
                            18 00 00 00 02 04 00 00 02 00 00 00 02 00 00 00
                            01 00 00 00 01 00 00 00 02 00 00 00 33 00 0B F0
                            12 00 00 00 BF 00 08 00 08 00 81 01 09 00 00 08
                            C0 01 40 00 00 08 40 00 1E F1 10 00 00 00 0D 00
                            00 08 0C 00 00 08 17 00 00 08 F7 00 00 10
              );

my $result     = $workbook->_add_mso_drawing_group();
   print $result;
