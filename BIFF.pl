use Spreadsheet::WriteExcel;


my $test_file   = 'temp_test_file.xls';
my $workbook    = Spreadsheet::WriteExcel->new($test_file);

my $worksheet  = $workbook->add_worksheet();
my @data = $worksheet->_comment_params(2, 0, 'Test');

my $row        = $data[0];
my $col        = $data[1];
my $author     = $data[4];
my $encoding   = $data[5];
my $visible    = $data[6];
my $obj_id     = 1;

my $result     = unpack_record($worksheet->_store_note($row,
                                                    $col,
                                                    $obj_id,
                                                    $author,
                                                    $encoding,
                                                    $visible,
                                                    ));
   print $result;


###############################################################################
#
# Unpack the binary data into a format suitable for printing in tests.
#
sub unpack_record {
    return join ' ', map {sprintf "%02X", $_} unpack "C*", $_[0];
}

