#!/usr/bin/perl -w

# T Giorgino May 2014 

# creates a Open Office Impress presentation from
# a list of images (assuming full page images)

# Inspired by a script by K.-H. Herrmann June 2005




use OpenOffice::OODoc;
use Getopt::Std;

our $opt_o="img2ooImpressExport.odp";
our $opt_a='';

getopts('o:a') or die "Parsing command line\n";

if (@ARGV == 0) {
  print << "EOF";
  Usage:
  img2ooImpress.pl [-a] [-o outputfile.odp] Image1.jpg Image2.jpg ...
     -a   Append to existing presentation (do not overwrite)
     -o   Output file
EOF

exit;
}

my $cont;
if($opt_a && -r $opt_o) {
	print "Appending\n";
	$cont=odfDocument(file=>$opt_o,part=>'content');
} else {
	print "Creating\n";
	$cont=odfDocument(file=>$opt_o,part=>'content',create=>'presentation');
}

$cont or die "Couldn't create document: $!\n";

my $i=1;
while (my $imgfile=shift @ARGV){
    my $test=$cont->insertDrawPage(-1);
    my $image=  $cont->createImageElement
	(
	 "Slide".$i." ".localtime,
	 description     => "image ".$i." filename:".$imgfile,
	 page            => $test,
	 position        => "0,0",
	 import          => $imgfile,
	 size            => "28cm, 21cm",
	 style           => "slide"
	);
    $i++;
}

$cont->save;

