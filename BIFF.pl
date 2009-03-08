use Spreadsheet::WriteExcel;


my $test_file   = 'temp_test_file.xls';
my $workbook    = Spreadsheet::WriteExcel->new($test_file);

my $worksheet  = $workbook->add_worksheet();

my @tests = (
    [
        undef,
        [],
    ],
);
for my $aref (@tests) {
    my $expression  = $aref->[0];
    my $expected    = $aref->[1];
    my @results     = $worksheet->_extract_filter_tokens($expression);

    my $testname    = $expression || 'none';


		print "result = $result\nexpected = $expected\n";
}


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


