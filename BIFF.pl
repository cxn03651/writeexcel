use Spreadsheet::WriteExcel;


my $test_file           = "temp_test_file.xls";
my $workbook            = Spreadsheet::WriteExcel->new($test_file);
my $worksheet           = $workbook->add_worksheet();
my $worksheet2          = $workbook->add_worksheet();
my $formula;
my $caption;
my @bytes;
my $target;
my $result;

$formula      = '=$E$3:$E$6';

$caption    = " \tData validation: _pack_dv_formula('$formula')";
@bytes      = qw(
                    09 00 0C 00 25 02 00 05 00 04 00 04 00
                );


# Zero out Excel's random unused word to allow comparison.
$bytes[2]   = '00';
$bytes[3]   = '00';
$target     = join " ", @bytes;

$result     = unpack_record($worksheet->_pack_dv_formula($formula));
		print "result = $result\nexpected = $expected\n";

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


