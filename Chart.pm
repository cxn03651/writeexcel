package Spreadsheet::WriteExcel::Chart;

###############################################################################
#
# Chart - A writer class for Excel Charts.
#
#
# Used in conjunction with Spreadsheet::WriteExcel
#
# Copyright 2000-2008, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

use Exporter;
use strict;
use Carp;
use FileHandle;
use Spreadsheet::WriteExcel::BIFFwriter;




use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::WriteExcel::BIFFwriter);

$VERSION = '2.22';

###############################################################################
#
# new()
#
# Constructor. Creates a new Chart object from a BIFFwriter object
#
sub new {

    my $class                   = shift;
    my $self                    = Spreadsheet::WriteExcel::BIFFwriter->new();

    $self->{_filename}          = $_[0];
    $self->{_name}              = $_[1];
    $self->{_index}             = $_[2];
    $self->{_encoding}          = $_[3];
    $self->{_activesheet}       = $_[4];
    $self->{_firstsheet}        = $_[5];

    $self->{_type}              = 0x0200;
    $self->{_ext_sheets}        = [];
    $self->{_using_tmpfile}     = 1;
    $self->{_filehandle}        = "";
    $self->{_fileclosed}        = 0;
    $self->{_offset}            = 0;
    $self->{_xls_rowmax}        = 0;
    $self->{_xls_colmax}        = 0;
    $self->{_xls_strmax}        = 0;
    $self->{_dim_rowmin}        = 0;
    $self->{_dim_rowmax}        = 0;
    $self->{_dim_colmin}        = 0;
    $self->{_dim_colmax}        = 0;
    $self->{_dim_changed}       = 0;
    $self->{_colinfo}           = [];
    $self->{_selection}         = [0, 0];
    $self->{_panes}             = [];
    $self->{_active_pane}       = 3;
    $self->{_frozen}            = 0;
    $self->{_selected}          = 0;
    $self->{_hidden}            = 0;

    $self->{_paper_size}        = 0x0;
    $self->{_orientation}       = 0x1;
    $self->{_header}            = '';
    $self->{_footer}            = '';
    $self->{_hcenter}           = 0;
    $self->{_vcenter}           = 0;
    $self->{_margin_head}       = 0.50;
    $self->{_margin_foot}       = 0.50;
    $self->{_margin_left}       = 0.75;
    $self->{_margin_right}      = 0.75;
    $self->{_margin_top}        = 1.00;
    $self->{_margin_bottom}     = 1.00;

    $self->{_title_rowmin}      = undef;
    $self->{_title_rowmax}      = undef;
    $self->{_title_colmin}      = undef;
    $self->{_title_colmax}      = undef;
    $self->{_print_rowmin}      = undef;
    $self->{_print_rowmax}      = undef;
    $self->{_print_colmin}      = undef;
    $self->{_print_colmax}      = undef;

    $self->{_print_gridlines}   = 1;
    $self->{_screen_gridlines}  = 1;
    $self->{_print_headers}     = 0;

    $self->{_fit_page}          = 0;
    $self->{_fit_width}         = 0;
    $self->{_fit_height}        = 0;

    $self->{_hbreaks}           = [];
    $self->{_vbreaks}           = [];

    $self->{_protect}           = 0;
    $self->{_password}          = undef;

    $self->{_col_sizes}         = {};
    $self->{_row_sizes}         = {};

    $self->{_col_formats}       = {};
    $self->{_row_formats}       = {};

    $self->{_zoom}              = 100;
    $self->{_print_scale}       = 100;

    $self->{_leading_zeros}     = 0;

    $self->{_outline_row_level} = 0;
    $self->{_outline_style}     = 0;
    $self->{_outline_below}     = 1;
    $self->{_outline_right}     = 1;
    $self->{_outline_on}        = 1;

    bless $self, $class;
    $self->_initialize();
    return $self;
}


###############################################################################
#
# _initialize()
#
sub _initialize {

    my $self       = shift;
    my $filename   = $self->{_filename};
    my $filehandle = FileHandle->new($filename) or
                     die "Couldn't open $filename in add_chart_ext(): $!.\n";

    binmode($filehandle);

    $self->{_filehandle} = $filehandle;
    $self->{_datasize}   = -s $filehandle;

}


###############################################################################
#
# _close()
#
# Add data to the beginning of the workbook (note the reverse order)
# and to the end of the workbook.
#
sub _close {

    my $self = shift;
}


###############################################################################
#
# get_name().
#
# Retrieve the worksheet name.
#
sub get_name {

    my $self    = shift;

    return $self->{_name};
}


###############################################################################
#
# get_data().
#
# Retrieves data from memory in one chunk, or from disk in $buffer
# sized chunks.
#
sub get_data {

    my $self   = shift;
    my $buffer = 4096;
    my $tmp;

    return $tmp if read($self->{_filehandle}, $tmp, $buffer);

    # No data to return
    return undef;
}


###############################################################################
#
# select()
#
# Set this worksheet as a selected worksheet, i.e. the worksheet has its tab
# highlighted.
#
sub select {

    my $self = shift;

    $self->{_hidden}         = 0; # Selected worksheet can't be hidden.
    $self->{_selected}       = 1;
}


###############################################################################
#
# activate()
#
# Set this worksheet as the active worksheet, i.e. the worksheet that is
# displayed when the workbook is opened. Also set it as selected.
#
sub activate {

    my $self = shift;

    $self->{_hidden}         = 0; # Active worksheet can't be hidden.
    $self->{_selected}       = 1;
    ${$self->{_activesheet}} = $self->{_index};
}


###############################################################################
#
# hide()
#
# Hide this worksheet.
#
sub hide {

    my $self = shift;

    $self->{_hidden}         = 1;

    # A hidden worksheet shouldn't be active or selected.
    $self->{_selected}       = 0;
    ${$self->{_activesheet}} = 0;
    ${$self->{_firstsheet}}  = 0;
}


###############################################################################
#
# set_first_sheet()
#
# Set this worksheet as the first visible sheet. This is necessary
# when there are a large number of worksheets and the activated
# worksheet is not visible on the screen.
#
sub set_first_sheet {

    my $self = shift;

    $self->{_hidden}         = 0; # Active worksheet can't be hidden.
    ${$self->{_firstsheet}}  = $self->{_index};
}




1;


__END__


=head1 NAME

Chart - A writer class for Excel Charts.

=head1 SYNOPSIS

See the documentation for Spreadsheet::WriteExcel

=head1 DESCRIPTION

This module is used in conjunction with Spreadsheet::WriteExcel.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

© MM-MMVIII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

