#!/usr/bin/perl -w

$^W = 1;
use strict;

use Getopt::Long;
use Docclient;
my $win_to_il2;
eval ' use Cz::Cstocs;
	$win_to_il2 = new Cz::Cstocs qw( 1250 il2 );
	';

use Text::Tabs;
my $tabstop = 8;	# implicit tab skip
my $WIDTH = 72;

my %options;
if ($0 =~ /^xls2/)
	{ $options{'in_format'} = 'xls'; }
if ($0 =~ /2html$/)
	{ $options{'out_format'} = 'html'; }

my $stdin = 0;
if (defined $ARGV[$#ARGV] and $ARGV[$#ARGV] eq '-')
	{ pop @ARGV; $stdin = 1; }

Getopt::Long::GetOptions( \%options,
	qw( raw help debug version server=s host=s port=s
		out_format=s in_format=s )
			) or exit;

if ($stdin or not @ARGV) { push @ARGV, '-' }

sub print_version
	{ print "This is docclient version $Docclient::VERSION.\n"; }

if (defined $options{'help'})
	{
	print_version();
	print <<"EOF";
usage: docclient [ options ] [ files ]
	where options is some of
		--host=host
		--server=host	name or address of your docserver
			(default $Docclient::Config::Config{'server'})
		--port=number	port number on your docserver
			(default $Docclient::Config::Config{'port'})
		--in_format=format	input format ($Docclient::Config::Config{'in_format'})
		--out_format=format	output format ($Docclient::Config::Config{'out_format'})
		--raw		do not clean the output in any way
		--version	version info
		--help		this help
	The format names depend on your server version.
EOF
	exit;
	}
if (defined $options{'version'})
	{ print_version(); exit; }
if (defined $options{'debug'})
	{ $Docclient::DEBUG = 1; }

print STDERR "Debug set to $Docclient::DEBUG\n" if $Docclient::DEBUG;

my $obj = new Docclient( %options ) or die $Docclient::errstr;

for my $file (@ARGV) {
	local *FILE;
	my $size;
	if ($file eq '-') {
		*FILE = \*STDIN;
	} else {
		open FILE, $file or die "Error reading $file: $!\n";
		$size = -s $file;
	}
	binmode FILE;
	$obj->put_file(*FILE, $file, $size);
	close FILE;

	$obj->convert($obj->{'in_format'}, $obj->{'out_format'}) or
		die "Error converting the data: ", $obj->errstr;

	if (not defined $obj->{'raw'})
		{
		if ($obj->{'out_format'} ne 'html')
			{
			# get all to scalar value, clean, print
			&clean_and_print_txt_data($obj->get_to_scalar());
			}
		else
			{
			print &clean_charset($obj->get_to_scalar());
			}
		}
	else {
		$obj->get_to_file(*STDOUT);
		}
	$obj->finished;
	}

sub clean_charset
	{
	if (defined $win_to_il2)
		{ return &$win_to_il2($_[0]); }
	$_[0];
	}

sub clean_and_print_txt_data
	{
	my $text = shift;

	# cancel spaces and LF at the end of lines
	$text =~ s/[ \t]*\r?\n/\n/g;

	# we do not want more than the subsequent empty lines
	$text =~ s/\n{3}\n+/\n\n\n/g;

	$text = &clean_charset($text);

	my $line;
	while ($text ne '' and $text =~ /(.*)\n/g)
		{
		my $line = $1;

		# expand tabulators
		$line = expand $line;

		# now try to fit the line into $WIDTH
		while ($line ne '')
			{
			my $length = length $line;
			if ($length <= $WIDTH)
				{ print $line; last; }

			# we try to compress spaces
			my @spaces = map { length $_ } $line =~ /(\s{2,})/g;
			my $sum = 0;
			for (@spaces) { $sum += $_; }
			my $shorting = $sum - scalar @spaces;

			if ($length - $shorting <= $WIDTH)
				{
				my $expand = 1 - ($length - $WIDTH) / $shorting;
				$line =~ s/(\s{2,})/ ' ' x ($expand * shift @spaces) /ge;
				print $line; last;	
				}

			my $start;
			$start = substr $line, 0, $WIDTH + 1;
			$start =~ s/(\b\w)?\s+\S+$//;
			print $start, "\n";
			$line = substr $line, length $start;
			$line =~ s/^\s+//;
			}
		print "\n";
		}
	}
1;
__END__

=head1 NAME

docclient.pm - client for remote conversions of MS format documents

=head1 SYNOPSIS

	docclient.pm msword.doc > out.txt
	docclient.pm --out_format=html msword.doc > out.html

=head1 AUTHOR

Michal Brandejs provided the code to clean up the txt/html on the
client side.

Jan Pazdziora did the original client/server implementation and the
basic Win32::OLE stuff.

Pavel Smerk added the code for other conversions (xls, HTML, prn,
cvs).

