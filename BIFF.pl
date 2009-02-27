use IO::File;
use Spreadsheet::WriteExcel::BIFFwriter;
use Spreadsheet::WriteExcel::Worksheet;
use Spreadsheet::WriteExcel::Format;

#my $ws = new Spreadsheet::WriteExcel::Worksheet('test', 0);
#my $fh;

#open ($fh, ">ws_store_filtermode_off");
#print {$fh} $ws->_store_filtermode;
#close $fh;

#$ws->autofilter(1,1,2,2);
#$ws->filter_column(1,'x < 2000');
#open ($fh, ">ws_store_filtermode_on");
#print {$fh} $ws->_store_filtermode;
#close $fh;


#	print $ws->_store_filtermode;


my $format = new Spreadsheet::WriteExcel::Format;

   print $format->set_font('Times New Roman');
