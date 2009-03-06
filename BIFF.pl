use Spreadsheet::WriteExcel;


my $test_file   = 'temp_test_file.xls';
my $workbook    = Spreadsheet::WriteExcel->new($test_file);
my $format      = $workbook->add_format();

my $worksheet  = $workbook->add_worksheet();
my $range   = 'A6';
   $worksheet->set_row(5,6);
   $worksheet->set_row(6,6);
   $worksheet->set_row(7,6);
   $worksheet->set_row(8,6);

my @data    = $worksheet->_substitute_cellref($range);
   @data    = $worksheet->_comment_params(@data, 'Test');
 	  @data    = @{$data[-1]};

my $result  = $worksheet->_store_mso_client_anchor(3, @data);
   print $result;
