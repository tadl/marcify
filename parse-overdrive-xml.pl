#!/usr/bin/perl
use strict;
use warnings;

use XML::LibXML;
use HTML::Parser;
use HTML::Entities;
use MARC::Record;
use MARC::File::USMARC;
use Business::ISBN;
use Config::Tiny;
use Getopt::Std;

use Data::Dumper;

# parse command line arguments
my %args;
getopts('vf:o:c:', \%args);

my $sourcefile = $args{'f'};
my $output_filename = $args{'o'} || "output.mrc";
my $config_file = $args{'c'} || "marcify.ini";
my $verbose = $args{'v'};


# Read and parse config file
my $Config = Config::Tiny->read( $config_file );
die "Failure reading config file $config_file\n" unless $Config;

my $lib_shortname = $Config->{"856"}->{"lib_shortname"} || die "no lib_shortname";
my $url_prefix = $Config->{"856"}->{"url_prefix"} || die "no url_prefix";
my $link_text = $Config->{"856"}->{"link_text"} || die "no link_text";

my @source_recs = &parse_file($sourcefile);

open my $marcfile, ">$output_filename";
binmode $marcfile, ":utf8"; # our records are UTF-8

map {
    my $marc = &make_marc($_);
    warn "writing marc record for ", $_->{"Title"}, "\n" if $verbose;
    print $marcfile MARC::File::USMARC->encode($marc);
} @source_recs;

close $marcfile;

sub make_marc {
    my $source = shift;
    my $marc = MARC::Record->new();
    $marc->leader('     nam a22     4u 4500');
                  #00056nam a22000374u 4500
                  #00129nam a22000614u 4500
    $marc->encoding( 'UTF-8' );
    my $field_008 = MARC::Field->new('008');
                        #0         1         2         3
                        #0123456789012345678901234567890123456789
    my $field_008_str = '      sDATE    CTY|||| o     ||| F|LNG d';

    # base DATE on DateOfPublication
    my $replace_date = &format_pubdate($source->{"DateOfPublication"});
    $field_008_str =~ s/DATE/$replace_date/;

    # base LNG on Language
    my $replace_lng = &get_lng_code($source->{"Language"});
    $field_008_str =~ s/LNG/$replace_lng/;

    # base CTY on PlaceOfPublication
    my $replace_cty = &get_cty_code($source);
    $field_008_str =~ s/CTY/$replace_cty/;

    # base F (literary form) on Subject containing Fiction or Nonfiction
    my $replace_litf = &get_litf_code($source->{"Subject"});
    $field_008_str =~ s/F/$replace_litf/;

    $field_008->update($field_008_str);
    $marc->append_fields($field_008);

    my $isbn = Business::ISBN->new($source->{"ISBN"});
    die("invalid isbn: " . $source->{"ISBN"}) unless $isbn->is_valid;

    &append_if_nonempty($marc, ['020',' ',' ','a'], $isbn->as_isbn13->as_string([]) . " (electronic bk. : Adobe EPUB)");
    &append_if_nonempty($marc, ['020',' ',' ','a'], $isbn->as_isbn10->as_string([]) . " (electronic bk. : Adobe EPUB)");

    &append_if_nonempty($marc, ['035',' ',' ','a'], "(OCoLC)" . $source->{"OclcControlNumber"}) if $source->{"OclcControlNumber"};
    &append_if_nonempty($marc, ['100','1',' ','a'], &format_name($source->{"Creator"}));

    my $nonfiling = &get_title_nonfiling($source->{"Title"});
    &append_subfields($marc, ['245','0',$nonfiling], ['a' => $source->{"Title"}, 'h' => '[electronic resource] /', 'c' => 'by ' . $source->{"Creator"} . '.']);

    &append_subfields($marc, ['260',' ',' '], ['a'=>$source->{"PlaceOfPublication"} . " :", 'b'=>$source->{"Publisher"} . ",", 'c'=>&format_pubdate($source->{"DateOfPublication"}) . "."]);
    &append_if_nonempty($marc, ['300',' ',' ','a'], '1 online resource (' . $source->{"FileSize"} . ' KB) :');
    &append_if_nonempty($marc, ['500',' ',' ','a'], "Description based on OverDrive metadata.");
    &append_if_nonempty($marc, ['520',' ',' ','a'], &format_description($source->{"ShortDescription"}));
    &append_subfields($marc, ['538',' ',' '], ['a' => $source->{"Format"} . " format, " . $source->{"SystemRequirements"} . " required to access."]);
    &append_subfields($marc, ['599', ' ', ' '], ['a' => "OVERDRIVE METADATA RECORD" ]);
    &append_subfields($marc, ['856','4','0'], ['9' => $lib_shortname, 'u' => $url_prefix . $source->{"URL"}, 'y' => $link_text]);

    return $marc;
}

sub get_lng_code {
    my $lang = shift;
    if ($lang eq 'English') {
        return 'eng';
    } else {
        die("Unknown language $lang");
    }
}

sub get_cty_code {
    my $source = shift;
    my $location = $source->{"PlaceOfPublication"};
    return "nyu" if $location eq 'New York';
    return "oru" if $location eq 'Ashland';
    return "ilu" if $location eq 'Carol Stream';
    return "miu" if $location eq 'Grand Rapids';
    die("Unknown location $location", Dumper($source));
}

sub get_litf_code {
    my $subject = shift;
    if ($subject =~ m/Fiction/) {
        return "1";
    } elsif ($subject =~ m/Nonfiction/ ) {
        return "0";
    } else {
        die("unable to determine literary form based on $subject");
    }
}

sub get_title_nonfiling {
    my $title = shift;
    #determine the length of common english article and any associated diacritics
    return length($1) if $title =~ m/^(['"]?(a|an|the)['"]? )/i;
    return 0;
}

sub format_pubdate {
#format publication date for 260$c
    my $pubdate_in = shift;
    return $1 if $pubdate_in =~ m@\d+/\d+/(\d\d\d\d)$@;
    return $1 if $pubdate_in =~ m@\w\w\w +\d+ (\d\d\d\d) 12:00AM$@;
    die("unable to extract year from publication date: $pubdate_in");
}

sub format_name {
#format Firstname Lastname as Lastname, Firstname. for 010 field
	my $input = shift;
	$input =~ m/(.*?) ((Van )?[^ ]+)(, Ph\.? ?D\.?)?$/;
	my $output = $2 . ", " . $1 . ".";
	$output =~ s/\.\.$/\./; # remove any trailing doubled dot
	return $output;
}

sub format_description {
#format ShortDescription for 520$a
    my $descin = shift;
    $descin = decode_entities($descin);
    my @text_array;
    my $parser = HTML::Parser->new(api_version => 3, 
                                    handlers => {
                                        text => [\@text_array, "text"],
                                    });
    $parser->parse($descin);
    $parser->eof();
    my $descout;
    my @desc_array;
    map {
        push(@desc_array, @$_[0]);
    } @text_array;

    $descout = join(" ", @desc_array); 

    return $descout;
}

sub append_if_nonempty {
    my $marc = shift;
    my $tag_array = shift;
    my $data = shift;

    if ($data && $data ne '') {
        $marc->append_fields(MARC::Field->new(@$tag_array => $data));
    } 
}

sub append_subfields {
    my $marc = shift;
    my $tag_array = shift;
    my $data = shift;

    $marc->append_fields(MARC::Field->new(@$tag_array => @$data));
}

sub parse_file {
    my $filename = shift;
    my $parser = XML::LibXML->new();

    open my $filehandle, "$filename";
    binmode $filehandle; # libxml2 handles encoding, so don't set :utf-8 here
    my $doc;
    eval {
        $doc = $parser->parse_fh($filehandle);
    };
    if ($@) {
        warn "first attempt to parse failed, falling back to Windows-1252 encoding\n" if $@;
        my $converted_string = &_return_win1252_file_as_utf8_string($filename);
        $doc = $parser->parse_string($converted_string);
    }

    my @sheets = $doc->documentElement()->getElementsByTagName("Worksheet");
    die "Didn't find exactly one sheet!\n" if (scalar(@sheets) != 1);

    my $sheet = $sheets[0];

    my @tables = $sheet->getElementsByTagName("Table");
    die "Didn't find exactly one table!\n" if (scalar(@tables) != 1);

    my $table = $tables[0];

    my @rows = $table->getElementsByTagName("Row");

    my $rowcount = 0;

    my $hh; # to lookup field names

    my @outrows;
    map {
        my $row = $_;

        if ($rowcount == 0) {
            $hh = _buildHeaderHash($row); #reference to header hash
        } else {

            my @cells = $row->getElementsByTagName("Cell");

            my $cell_count = 0;
            my %rowhash;
            map {
                my $data = $_->getElementsByTagName("Data")->[0]->textContent();
                my $dataname = $hh->{$cell_count};
                $data = decode_entities($data); # ShortDescription and Subject have html entities
                $rowhash{$dataname} = $data;
                $cell_count++;
            } @cells;
            push(@outrows, \%rowhash);
        }

        $rowcount++;

    } @rows;

    return @outrows;
}

sub _buildHeaderHash {
    my $header_row = shift;
    # walk cells and build a position->name mapping
    my @header_cells = $header_row->getElementsByTagName("Cell");
    my %header_hash;
    my $header_count = 0;
    map {
        my $data = $_->getElementsByTagName("Data")->[0]->textContent();
        $header_hash{$header_count} = $data;
        $header_count++;
    } @header_cells; 

    return \%header_hash; 
}

sub _return_win1252_file_as_utf8_string {
    use Encoding;

    my $infile = shift;

    # open input file in raw mode
    open my $in, "$infile";
    binmode $in;

    # read entire file into string
    undef $/;
    my $infile_string = <$in>;

    # convert from Windows-1252 to utf8
    Encode::from_to($infile_string, 'Windows-1252', 'utf8');

    return $infile_string;
}
