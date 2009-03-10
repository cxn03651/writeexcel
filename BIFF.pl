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

$workbook->{_biff_only} = 1;

# store_workbook();

# store_worksheet_part_1();

#print '$worksheet->_store_window2();'."\n";
#$worksheet->_store_window2();
#$data = unpack_record($worksheet->{_data});
#print "$data\n";




#print '$workgook->get_data();'."\n";
#$data = unpack_record($workbook->get_data());
#print "$data\n";

#exit;


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


sub store_workbook_part_1 {
my $data;
$data = unpack_record($workbook->{_data});
print "$data\n";

    $workbook->_store_bof(0x0005);
print "    _store_bof(0x0005);\n";
$data = unpack_record($workbook->{_data});
print "$data\n";

    $workbook->_store_codepage();
print "    _store_codepage();\n";
$data = unpack_record($workbook->{_data});
print "$data\n";

    $workbook->_store_window1();
print "    _store_window1();\n";
$data = unpack_record($workbook->{_data});
print "$data\n";

    $workbook->_store_hideobj();
print "    _store_hideobj();\n";
$data = unpack_record($workbook->{_data});
print "$data\n";

    $workbook->_store_1904();
print "    _store_1904();\n";
$data = unpack_record($workbook->{_data});
print "$data\n";

    $workbook->_store_all_fonts();
print "    _store_all_fonts();\n";
$data = unpack_record($workbook->{_data});
print "$data\n";

    $workbook->_store_all_num_formats();
print "    _store_all_num_formats();\n";
$data = unpack_record($workbook->{_data});
print "$data\n";

    $workbook->_store_all_xfs();
print "    _store_all_xfs();\n";
$data = unpack_record($workbook->{_data});
print "$data\n";

    $workbook->_store_all_styles();
print "    _store_all_styles();\n";
$data = unpack_record($workbook->{_data});
print "$data\n";

    $workbook->_store_palette();
print "    _store_palette();\n";
$data = unpack_record($workbook->{_data});
print "$data\n";

$workbook->_store_boundsheet($worksheet->{_name},
                                 $worksheet->{_offset},
                                 $worksheet->{_type},
                                 $worksheet->{_hidden},
                                 $worksheet->{_encoding});
print "_store_boundsheet\n";
$data = unpack_record($workbook->{_data});
print "$data\n";


}

sub store_worksheet_part_1 {
my $data;
$data = unpack_record($worksheet->{_data});
print "$data\n";


print '$worksheet->_store_dimensions();'."\n";
$worksheet->_store_dimensions();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_autofilters();'."\n";
$worksheet->_store_autofilters();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_autofilterinfo();'."\n";
$worksheet->_store_autofilterinfo();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_filtermode();'."\n";
$worksheet->_store_filtermode();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_defcol();'."\n";
$worksheet->_store_defcol();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_password();'."\n";
$worksheet->_store_password();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_protect();'."\n";
$worksheet->_store_protect();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_obj_protect();'."\n";
$worksheet->_store_obj_protect();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_setup();'."\n";
$worksheet->_store_setup();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_margin_bottom();'."\n";
$worksheet->_store_margin_bottom();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_margin_top();'."\n";
$worksheet->_store_margin_top();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_margin_right();'."\n";
$worksheet->_store_margin_right();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_margin_left();'."\n";
$worksheet->_store_margin_left();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_vcenter();'."\n";
$worksheet->_store_vcenter();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_hcenter();'."\n";
$worksheet->_store_hcenter();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_footer();'."\n";
$worksheet->_store_footer();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_header();'."\n";
$worksheet->_store_header();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_vbreak();'."\n";
$worksheet->_store_vbreak();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_hbreak();'."\n";
$worksheet->_store_hbreak();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_wsbool();'."\n";
$worksheet->_store_wsbool();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_defrow();'."\n";
$worksheet->_store_defrow();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_guts();'."\n";
$worksheet->_store_guts();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_gridset();'."\n";
$worksheet->_store_gridset();
$data = unpack_record($worksheet->{_data});
print "$data\n";
print '$worksheet->_store_print_gridlines();'."\n";
$worksheet->_store_print_gridlines();
$data = unpack_record($worksheet->{_data});
print "$data\n";

print '$worksheet->_store_print_headers();'."\n";
$worksheet->_store_print_headers();
$data = unpack_record($worksheet->{_data});
print "$data\n";

}
