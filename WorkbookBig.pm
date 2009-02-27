package Spreadsheet::WriteExcel::WorkbookBig;

###############################################################################
#
# WorkbookBig - A writer class for Excel Workbooks > 7MB.
#
#
# Used in conjunction with Spreadsheet::WriteExcel
#
# Copyright 2000-2008, John McNamara and Kawai Takanori.
#
# Documentation after __END__
#

use Exporter;
use strict;
use Carp;
use Spreadsheet::WriteExcel::Workbook;
use OLE::Storage_Lite;
use Spreadsheet::WriteExcel::Worksheet;
use Spreadsheet::WriteExcel::Format;


use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::WriteExcel::Workbook Exporter);

$VERSION = '2.22';

###############################################################################
#
# new()
#
# Constructor. Creates a new WorkbookBig object from a Workbook object.
#
sub new {

    my $class = shift;
    my $self  = Spreadsheet::WriteExcel::Workbook->new(@_);

    # Drop some compatibility to save memory and speed up big files.
    $self->{_compatibility} = 0;

    bless $self, $class;
    return $self;
}


###############################################################################
#
# _store_OLE_file(). Over-ridden.
#
# Store the workbook in an OLE container using OLE::Storage_Lite.
#
sub _store_OLE_file {

    my $self = shift;

    my $stream  = pack 'v*', unpack 'C*', 'Workbook';
    my $OLE     = OLE::Storage_Lite::PPS::File->newFile($stream);


    my $tmp;
    $OLE->append($tmp) while $tmp = $self->get_data();

    foreach my $worksheet (@{$self->{_worksheets}}) {
        $OLE->append($tmp) while $tmp = $worksheet->get_data();
    }

    my @ltime = localtime();
    splice(@ltime, 6);
    my $date = OLE::Storage_Lite::PPS::Root->new(\@ltime, \@ltime,[$OLE]);
    $date->save($self->{_filename});

    # Close the filehandle if it was created internally.
    return CORE::close($self->{_fh_out}) if $self->{_internal_fh};


}


1;


__END__


=head1 NAME

WorkbookBig - A writer class for Excel Workbooks > 7MB.


=head1 SYNOPSIS

The direct use of this module is deprecated. See below.


=head1 DESCRIPTION

The module is a sub-class of Spreadsheet::WriteExcel used for creating Excel files greater than 7MB.

Direct use of this module is deprecated. As of version 2.17 Spreadsheet::WriteExcel can create files larger than 7MB if OLE::Storage_Lite is installed.

This module only exists for backwards compatibility.


=head1 REQUIREMENTS

OLE::Storage_Lite


=head1 AUTHORS

John McNamara jmcnamara@cpan.org

Kawai Takanori kwitknr@cpan.org


=head1 COPYRIGHT

© MM-MMVIII, John McNamara and Kawai Takanori.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
