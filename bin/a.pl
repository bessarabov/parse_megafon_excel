#!/usr/bin/perl

=encoding UTF-8
=cut

=head1 DESCRIPTION

=cut

# common modules
use strict;
use warnings FATAL => 'all';
use feature 'say';
use utf8;
use open qw(:std :utf8);

use DDP;
use Carp;
use XML::Simple;

# main
sub main {

    my $ref = XMLin('/data/a.xlsx');

# $ref is a hashref with keys:
#[
#    [0] "Worksheet",
#    [1] "xmlns:o",
#    [2] "xmlns:html",
#    [3] "Styles",
#    [4] "ExcelWorkbook",
#    [5] "xmlns:x",
#    [6] "DocumentProperties",
#    [7] "xmlns:ss",
#    [8] "xmlns"
#]

# $ref->{Worksheet} is a hashref with keys:
#[
#    [0] "Table",
#    [1] "PageBreaks",
#    [2] "ss:Name",
#    [3] "WorksheetOptions"
#]

# $ref->{Worksheet}->{Table} is a hashref with keys:
#[
#    [0] "x:FullColumns",
#    [1] "Column",
#    [2] "ss:ExpandedRowCount",
#    [3] "ss:ExpandedColumnCount",
#    [4] "x:FullRows",
#    [5] "Row"
#]

# $ref->{Worksheet}->{Table}->{Column} is a arrayref;
#\ [
#    [0]  {
#        ss:AutoFitWidth   0,
#        ss:Width          8.24
#    },
#    [1]  {
#        ss:AutoFitWidth   0,
#        ss:Width          12.36
#    },

# $ref->{Worksheet}->{Table}->{Row} is a arrayref;

    foreach my $row (@{$ref->{Worksheet}->{Table}->{Row}}) {

        my @data;

        foreach my $element (@{$row->{Cell}}) {
            if ($element->{Data}) {
                push
                    @data,
                    $element->{Data}->{content} // ''
                    ;
            } else {
                push
                    @data,
                    ''
                    ;
            }
        }

        print join "\t", @data;
        print "\n";
    }
}
main();
__END__
