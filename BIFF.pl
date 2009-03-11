use Spreadsheet::WriteExcel;


my $test_file   = "temp_test_file.xls";
my $workbook    = Spreadsheet::WriteExcel->new($test_file);
my $worksheet;
my @tests, @data;

$workbook->compatibility_mode(1);

# Test 1.
#
$worksheet = $workbook->add_worksheet();
$range   = 'A6';
$worksheet->set_row(5,6);
$worksheet->set_row(6,6);
$worksheet->set_row(7,6);
$worksheet->set_row(8,6);
@data    = $worksheet->_substitute_cellref($range);
@data    = $worksheet->_comment_params(@data, 'Test');
@data    = @{$data[-1]};

$workbook->{_biff_only} = 1;

$workbook->close();

#
# Helper functions.
#

###############################################################################
#
# Unpack the binary data into a format suitable for printing in tests.
#
sub unpack_record {
    return join ' ', map {sprintf "%02X", $_} unpack "C*", $_[0];
}


