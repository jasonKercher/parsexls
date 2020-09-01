#!/usr/bin/perl

use strict;
use warnings;

#Including libraries from /prog/lib/perl_lib
use lib '/prog/lib/perl_lib/lib/perl5/5.18.2/x86_64-linux-thread-multi';
use lib '/prog/lib/perl_lib/lib/perl5/5.18.2';
use lib '/prog/lib/perl_lib/lib/perl5/x86_64-linux-thread-multi';
use lib '/prog/lib/perl_lib/lib/perl5';
use lib '/prog/lib/perl_lib/lib/perl5/5.18.0/x86_64-linux-thread-multi';
use lib '/prog/lib/perl_lib/lib/perl5/5.18.0';
use lib '/prog/lib/perl_lib/lib/perl5/5.18.1/x86_64-linux-thread-multi';
use lib '/prog/lib/perl_lib/lib/perl5/5.18.1 /usr/lib/perl5/site_perl';

use Getopt::Long;
use File::Type;
use File::Basename;
#use Spreadsheet::WriteExcel;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseXLSX;

# Variables for command line arguments
my $select_sheet_number;
my $select_sheet;
my $output_file;
my $active_text_qualifiers;
my $rfc4180;
my $expand_merged;
my $multiple_csvs;
my $output_unformatted;
my $print_help;
my $delm = "|";

# Pull options from @ARGV. -s for sheet name/number or -x for expanding merged values
Getopt::Long::GetOptions(
'n=i' => \$select_sheet_number,
's=s' => \$select_sheet,
'd=s' => \$delm,
'o=s' => \$output_file,
'q'   => \$active_text_qualifiers,
'Q'   => \$rfc4180,
'u'   => \$output_unformatted,
'h'   => \$print_help,
'x'   => \$expand_merged,
'm'   => \$multiple_csvs);

if( $print_help ) {
  print "
Usage: $0 <source file>

  Options:
    -n ... Select a sheet by Number.
    -s ... Select a Sheet by name (regex may be used here).
    -o ... Name the output file. Overridden by -m.
    -m ... Output a csv for each sheet (named for the sheet).
    -x ... eXpand all merged cell values across merged area.
    -q ... Activate text Qualifiers in the output.
    -Q ... Activate RFC-4180 text qualifiers.
    -d ... Set an output Delimiter. Default is \"|\" (pipe).
    -u ... Output the Unformatted (true) value in each cell.
    -h ... Print this help menu.\n\n";
  exit;
}

# text qualifiers are implied with rfc4180 option
if ($rfc4180) {
  $active_text_qualifiers = $rfc4180;
}

# Parse input file name
my $clientFile = shift @ARGV or die "usage: $0 <source file> (-h for help)\n";
my ($csvFile, $csvDir, $csvExt) = fileparse($clientFile, qr/\.[^.]*/);
my $clientCsvFile = "${csvDir}${csvFile}.csv";
my ($clientXls, $clientBook);

# Check file type: if it is of type zip, it could be in xlsx format
my $ft = File::Type->new();
if ( $ft->mime_type($clientFile) eq 'application/zip' ) {
  $clientXls = new Spreadsheet::ParseXLSX;
  $clientBook = $clientXls->parse($clientFile) or die "Could not open Client Excel file $clientFile: $!";
}
else {
  $clientXls = new Spreadsheet::ParseExcel;
  $clientBook = $clientXls->Parse($clientFile) or die "Could not open Client Excel file $clientFile: $!";
}

# Open file for output
my $last_loop = 0;
my $fh;
if ($multiple_csvs) {
  $last_loop = 2;
}
else {
  unless(defined $select_sheet_number or defined $select_sheet or defined $output_file) {
    open($fh, '>', $clientCsvFile) or die "Could not open file $clientCsvFile: $!";
  }
  if(defined $output_file) {
    $clientCsvFile = $output_file;
    open($fh, '>', $clientCsvFile) or die "Could not open file $clientCsvFile: $!";
  }
}


# Loop through sheets
foreach my $clientSheetNumber (0 .. $clientBook->{SheetCount}-1) {
  my $source_sheet;

  # Select Sheet by number
  if(defined $select_sheet_number) {
    if ( $select_sheet_number > $clientBook->{SheetCount} or $select_sheet_number < 1 ) {
      die "The number you selected is not an available sheet!";
    }
    else {
      $clientSheetNumber = $select_sheet_number - 1;
      $last_loop = 1;
    }
  }

  $source_sheet = $clientBook->{Worksheet}[$clientSheetNumber];

  # Select Sheet by name
  if(defined $select_sheet) {
    #my $re = qr/$select_sheet/;
    unless( lc($source_sheet->{Name}) =~ /$select_sheet/i ) {
      next;
    }
    $last_loop = 1;
  }
  elsif (not $last_loop) {
    print $fh "--------- SHEET:", $source_sheet->{Name}, "\n";
  }

  if ($multiple_csvs or ($last_loop and not defined $output_file)) {
    $clientCsvFile = $source_sheet->{Name} . ".csv";
    open($fh, '>', $clientCsvFile) or die "Could not open file $clientCsvFile: $!";
  }

  next unless defined $source_sheet->{MaxRow};
  next unless $source_sheet->{MinRow} <= $source_sheet->{MaxRow};
  next unless defined $source_sheet->{MaxCol};
  next unless $source_sheet->{MinCol} <= $source_sheet->{MaxCol};
  my @table;

  # Read Data from spreadsheet into @table
  foreach my $row_index ($source_sheet->{MinRow} .. $source_sheet->{MaxRow}) {

    my @row;
    foreach my $col_index ($source_sheet->{MinCol} .. $source_sheet->{MaxCol}) {
      my $source_cell = $source_sheet->{Cells}[$row_index][$col_index];

      if ($source_cell) {
        if ($output_unformatted) {
          $row[$col_index] = $source_cell->unformatted();
        }
        else {
          $row[$col_index] = $source_cell->value();
        }
      }
    }
    $table[$row_index] = \@row;
  }

  # If -x option given, copy top left cell value accross all merged cells.
  if($expand_merged) {
    my $merged_areas = $source_sheet->get_merged_areas();
    for my $merge_group ( @$merged_areas ) {
      if( $merge_group ) {
        my $hold_value = $table[ $merge_group->[0] ][ $merge_group->[1] ];
        for my $row_index ( $merge_group->[0] .. $merge_group->[2] ) {
          for my $col_index ( $merge_group->[1] .. $merge_group->[3] ) {
            $table[$row_index][$col_index] = $hold_value;
          }
        }
      }
    }
  }

  # Spit out array of arrays (@table) to file.
  for my $i ( 0 .. $#table ) {
    if( $table[$i] ) {
      my @row = @{ $table[$i] };

      for my $j ( 0 .. $#row ) {

        if ( defined $row[$j] ) {
          if ( $active_text_qualifiers and ($row[$j] =~ /\Q$delm\E/ or $row[$j] =~ /\"/ or $row[$j] =~ /\n/) ) {
            if ($rfc4180) {
              $row[$j] =~ s/\"/\"\"/g;
            }
            print $fh "\"$row[$j]\"";
          }
          else {
            print $fh $row[$j];
          }
        }

        if( $j eq $#row ) {
          print $fh "\n";
        }
        else {
          print $fh $delm;
        }
      }
    }
  }

  if( $multiple_csvs ) {
    print "$clientCsvFile\n";
    close $fh;
  }

  if( $last_loop == 1 ) {
    last;
  }
}
### END convert to CSV  ##
#
unless( $multiple_csvs ) {
  print "$clientCsvFile\n";
}

