use Spreadsheet::WriteExcel;


my $test_file   = "temp_test_file.xls";
my $workbook    = Spreadsheet::WriteExcel->new($test_file);
my $worksheet;
my @rows;
my @tests;
my $row;
my $col1;
my $col2;

$workbook->compatibility_mode(1);

# Test 1.
#
$row  = 1;
$col1 = 0;
$col2 = 0;
$worksheet = $workbook->add_worksheet();
$worksheet->set_row($row, 15);
push @tests,    [
                    " \tset_row(): row = $row, col1 = $col1, col2 = $col2",
                    {
                        col_min => 0,
                        col_max => 0,
                    }
                ];


# Test 2.
#
$row  = 2;
$col1 = 0;
$col2 = 0;
$worksheet = $workbook->add_worksheet();
$worksheet->write($row, $col1, 'Test');
$worksheet->write($row, $col2, 'Test');
push @tests,    [
                    " \twrite():   row = $row, col1 = $col1, col2 = $col2",
                    {
                        col_min => 0,
                        col_max => 1,
                    }
                ];

$workbook->{_biff_only} = 1;

$workbook->close();

open    XLSFILE, $test_file or die "Couldn't open test file\n";
binmode XLSFILE;

my $header;
my $data;
while (read XLSFILE, $header, 4) {

    my ($record, $length) = unpack 'vv', $header;
    read XLSFILE, $data, $length;

    # Read the row records only.
    next unless $record == 0x0208;
    my ($col_min, $col_max) = unpack 'x2 vv', $data;

    push @rows,
                {
                    col_min => $col_min,
                    col_max => $col_max,
                };
}


for my $i (0 .. @tests -1) {

#    is_deeply($rows[$i], $tests[$i]->[1], $tests[$i]->[0]);
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


