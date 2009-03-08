use Spreadsheet::WriteExcel;


$testname = '3w x 5h jpeg image.';
$image    = pack 'H*', join '', qw (

    FF D8 FF E0 00 10 4A 46 49 46 00 01 01 01 00 60
    00 60 00 00 FF DB 00 43 00 06 04 05 06 05 04 06
    06 05 06 07 07 06 08 0A 10 0A 0A 09 09 0A 14 0E
    0F 0C 10 17 14 18 18 17 14 16 16 1A 1D 25 1F 1A
    1B 23 1C 16 16 20 2C 20 23 26 27 29 2A 29 19 1F
    2D 30 2D 28 30 25 28 29 28 FF DB 00 43 01 07 07
    07 0A 08 0A 13 0A 0A 13 28 1A 16 1A 28 28 28 28
    28 28 28 28 28 28 28 28 28 28 28 28 28 28 28 28
    28 28 28 28 28 28 28 28 28 28 28 28 28 28 28 28
    28 28 28 28 28 28 28 28 28 28 28 28 28 28 FF C0
    00 11 08 00 05 00 03 03 01 22 00 02 11 01 03 11
    01 FF C4 00 15 00 01 01 00 00 00 00 00 00 00 00
    00 00 00 00 00 00 00 07 FF C4 00 14 10 01 00 00
    00 00 00 00 00 00 00 00 00 00 00 00 00 00 FF C4
    00 15 01 01 01 00 00 00 00 00 00 00 00 00 00 00
    00 00 00 06 08 FF C4 00 14 11 01 00 00 00 00 00
    00 00 00 00 00 00 00 00 00 00 00 FF DA 00 0C 03
    01 00 02 11 03 11 00 3F 00 9D 00 1C A4 5F FF D9
);

@expected = ($type, 3, 5);
@results  = Spreadsheet::WriteExcel::Workbook::_process_jpg(1, $image, 'test.jpg');
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


