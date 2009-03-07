use Spreadsheet::WriteExcel;


my $test_file   = 'temp_test_file.xls';
my $workbook    = Spreadsheet::WriteExcel->new($test_file);

my $worksheet  = $workbook->add_worksheet();

my @tests = (
    {
        'column'        => 0,
        'expression'    => 'x =  Blanks',
        'data'          => [qw(
                                9E 00 18 00 00 00 84 32 0C 02 00 00 00 00 00 00
                                00 00 00 00 00 00 00 00 00 00 00 00

                           )],
    },
);
for my $test (@tests) {

    my $column     = $test->{column};
    my $expression = $test->{expression};
    my @tokens     = $worksheet->_extract_filter_tokens($expression);
       @tokens     = $worksheet->_parse_filter_expression($expression, @tokens);

    my $result = $worksheet->_store_autofilter($column , @tokens);

    my $target     = join " ",  @{$test->{data}};

    my $caption    = " \tfilter_column($column, '$expression')";

    $result     = unpack_record($result);
    is($result, $target, $caption);
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


