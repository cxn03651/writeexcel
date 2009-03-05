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
my $formula = $worksheet->store_formula('=A1 * 3 + 50');
$worksheet->repeat_formula(5, 3, $formula, $format, 'A1', 'A2');

$data               = $worksheet ->_store_dimensions();
@results {@dims}    = unpack 'x4 VVvv', $data;
@expected{@dims}    = (0, 1, 0, 1);
