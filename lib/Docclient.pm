
package Docclient;

use IO::Handle;
use RPC::PlClient;
use Docclient::Config;

use strict;
use vars qw( $VERSION $errstr $DEBUG );

$VERSION = '0.85';

$DEBUG = 0;
sub debug ($) {
	return unless $DEBUG;
	my $txt = shift; $txt =~ s/([^\n])$/$1\n/; print STDERR $txt;
	}

sub new
	{
	my $class = shift;
	my %options = ( %Docclient::Config::Config, @_ );

	my %serveroptions;
	$serveroptions{'peeraddr'} = ( $options{'host'} or $options{'server'});
	$serveroptions{'peerport'} = $options{'port'};
	$serveroptions{'application'} = $options{'ServerClass'};
	$serveroptions{'version'} = 0.81;

	my $self;

	eval {
		debug "Connecting to server $serveroptions{'peeraddr'}:$serveroptions{'peerport'}";
		my $client = eval { RPC::PlClient->new( %serveroptions ); } or die $@;

		debug 'Requesting Docserver';
		my $obj = eval { $client->ClientObject('Docserver', 'new'); };
		if ($@) {
			my ($stderr) = $client->Call('errstr');
			die $stderr;
			}
		debug 'Negotiating chunk size';
		my $ChunkSize = $obj->preferred_chunk_size($options{'ChunkSize'});
		$self = bless {
			%options,
			'obj' => $obj,		
			'ChunkSize' => $ChunkSize,
			}, $class;
		};
	if ($@)
		{ $errstr = $@; return; }
	$self;
	}

sub put_file
	{
	my ($self, $fh, $file, $size) = @_;
	my $obj = $self->{'obj'};

	if (not defined $size) { $size = -1; }

	eval {
		debug "Processing $file (size $size)";
		$obj->input_file_length($size);

		my $buflen = $self->{'ChunkSize'};
		debug "Setting chunk size to $buflen";
		my $written = 0;
		while ($size < 0 or $written < $size)
			{
			my $buffer;
			my $out = $fh->read($buffer, $buflen);
			if ($out == 0)
				{
				debug "Strange: read returned 0 after reading $written bytes\n";
				last;
				}
			$written += $out;

			$obj->put($buffer);
			}
		debug "Written $written bytes";	
		return 1;
		};

	if ($@)
		{ $self->{'errstr'} = "Error occured: $@"; }
	return;
	}

sub put_scalar
	{
	my ($self, $data) = @_;
	my $obj = $self->{'obj'};
	my $size = length $data;

	eval {
		debug "Processing scalar data (size $size)";
		$obj->input_file_length($size);
		$obj->put($data);
		return 1;
		};

	if ($@)
		{ $self->{'errstr'} = "Error occured: $@"; }
	return;
	}

sub convert
	{
	my ($self, $in_format, $out_format) = @_;
	my $obj = $self->{'obj'};
	debug "Calling convert($in_format, $out_format)";
	$obj->convert($in_format, $out_format)
		or do { $self->{'errstr'} = $obj->errstr; return; };
	1;
	}

sub get_to_file
	{
	my ($self, $fh) = @_;
	debug "Calling get_to_file($fh)";
	my $obj = $self->{'obj'};
	my $buflen = $self->{'ChunkSize'};

	my $result_length = $obj->result_length;
	debug "Result length is $result_length\n";

	my $read = 0;
	while ($read < $result_length) {
		my $buffer = $obj->get($buflen);
		if (length $buffer == 0)
			{
			debug "Strange: read returned 0 after reading $read bytes\n";
			last;
			}
		$read += length $buffer;
		$fh->print($buffer);
		}
	1;
	}

sub get_to_scalar
	{
	my ($self, $fh) = @_;
	my $obj = $self->{'obj'};
	my $buflen = $self->{'ChunkSize'};

	my $result_length = $obj->result_length;
	debug "Result length is $result_length\n";

	my $result = '';
	my $read = 0;
	while ($read < $result_length) {
		my $buffer = $obj->get($buflen);
		if (length $buffer == 0)
			{
			debug "Strange: read returned 0 after reading $read bytes\n";
			last;
			}
		$read += length $buffer;
		$result .= $buffer;
		}
	$result;
	}

sub finished
	{
	my $self = shift;
	my $obj = $self->{'obj'};
	$obj->finished;
	}
sub DESTROY
	{
	shift->finished;
	}

1;

=head1 NAME

Docclient - client module for MS format conversions

=head1 SEE ALSO

docclient(1), Docserver(3), Win32::OLE(3)

=cut


