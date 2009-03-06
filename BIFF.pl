use Spreadsheet::WriteExcel;


my $test_file   = 'temp_test_file.xls';
my $workbook    = Spreadsheet::WriteExcel->new($test_file);
my $format      = $workbook->add_format();
my $worksheet;
my @dims        = qw(row_min row_max col_min col_max);
my $data;
my $caption;
my %results;
my %expected;
my $error;
my $smiley = pack "n", 0x263a;

$worksheet  = $workbook->add_worksheet();
my $result = $worksheet->convert_date_time('1900-01-01T');
print $result;
