
package Docserver;
use strict;
use IO::File;
use Fcntl;
use Win32::OLE;
use Win32::OLE::Const 'Microsoft Word';
use Win32::OLE::Const 'Microsoft Excel';

$Docserver::VERSION = 0.81;

# Placeholder for global error string
use vars qw( $errstr );

# Values for output formats
my %docoutform = (
		'txt' =>	wdFormatText, 
###		'txt' =>	9,
		'txt1' =>	wdFormatTextLineBreaks,
		'rtf' =>	wdFormatRTF,
		'doc6' =>	12,
		'doc95' =>	12,
		'html' =>	11
		); 
my %xlsoutform = (
		'prn' =>	xlTextPrinter,
		'csv' =>	xlCSV,
###		'txt' =>	xlTextPrinter, 
		'xls5' =>	xlExcel5,
		'xls95' =>	xlExcel5
		);
my $SHEET_SEP = "\n##### List è. %s #####\n\n";

# Create Docserver object, initialize input temporary file
sub new
	{
	my $class = shift;
	my $self;
	eval {
		$self = bless {}, $class;

		my ($dir, $filename, $fill, $ext) = ("c:/tmp/", "docserver${$}x", '', '.lck');
		if (not -d $dir)
			{ die "Directory $dir doesn't exist\n"; }
		$filename = $dir.$filename;
		my $fh;
		while (($fill eq '' or $fill < 1000) and
#			not $fh->open('>'.$filename.$fill.$ext))
#			not defined($fh = new IO::File '>'.$filename.$fill.$ext,
#							O_EXCL|O_WRONLY|O_binmode))
			( -f $filename.$fill.$ext
				or not defined($fh = new IO::File '>'.$filename.$fill.$ext)))
			{ $fill++; }
		if (not defined $fh)
			{ die "Couldn't create tmp file, even if I tried\nupto $filename$fill$ext: $@\n"; }
		$self->{'lockfile'} = $filename.$fill.$ext;
	
		( $self->{'infile'} = $self->{'lockfile'} ) =~ s/\.lck$/.in/;

#		$self->{'fh'} = new IO::File $self->{'infile'}, O_EXCL|O_WRONLY|O_binmode
		$self->{'fh'} = new IO::File ">$self->{'infile'}"
			or die "Error: $!\n";
		binmode $self->{'fh'};
		};
	if ($@)
		{
		### print STDERR "Whao, some error: $@ at line ", __LINE__, "\n";
		$errstr = $@; return;
		}
	### print STDERR "No error, returning $self with $self->{'fh'} at line ", __LINE__, "\n";
	$self;
	}
sub errstr
	{
	my $ref = shift;
	if (defined $ref and ref $ref) { return $ref->{'errstr'}; }
	return $errstr;
	}
sub preferred_chunk_size
	{
	my ($self, $size) = @_;
	$size;
	}
sub input_file_length
	{
	my $self = shift;
	$self->{'input_file_length'} = shift;
	}
sub put
	{
	my $self = shift;
	print { $self->{'fh'} } shift;
	1;
	}
sub convert
	{
	my ($self, $in_format, $out_format) = @_;
	delete $self->{'errstr'};

##	print STDERR "Called convert at line ", __LINE__, "\n";

	eval {
		if (defined $self->{'fh'})
			{
			$self->{'fh'}->close();
			delete $self->{'fh'};
			}
		if ($in_format eq 'doc')
			{
			if (not defined $docoutform{$out_format})
				{ die "Unknown output format $out_format"; }
			$self->doc_convert($out_format);
			}
		elsif ($in_format eq 'xls')
			{
			if ($out_format eq 'txt')
				{ $out_format = 'prn'; }
			if (not defined $xlsoutform{$out_format})
				{ die "Unknown output format $out_format"; }
			$self->xls_convert($out_format);
			}
		else
			{ die "Unknown input format $in_format"; }
		};
	if ($@)
		{
		$self->{'errstr'} = $@;
		print STDERR "Setting errstr to problem: $@\n";
		}
	1 unless defined $self->{'errstr'};
	}

sub doc_convert
	{
	my ($self, $out_format) = @_;
	my $newname;
	($newname = $self->{'infile'}) =~ s/\.in$/.doc/;
	rename $self->{'infile'}, $newname
		or die "Error moving $self->{'infile'} to $newname";
print STDERR "Newname $newname\n";
	my $word = ( Win32::OLE->GetActiveObject('Word.Application')
		or Win32::OLE->new('Word.Application', 'Quit') )
		or die Win32::OLE->LastError;
	my $doc = $word->Documents->Open($newname)
				or die Win32::OLE->LastError;
	if ($out_format eq 'ps') {
		($self->{'outfile'} = $self->{'infile'}) =~ s/\.in$/.prn/;
		my $origback = $word->Options->{PrintBackground};
		$word->Options->{PrintBackground} = 0;
		$doc->Activate;
		$word->PrintOut( {
			'Range' => wdPrintAllDocument,
			'PrintToFile' => 1, 
			'OutputFileName' => $self->{'outfile'},
			'Copies' => 1
			} );
		rename $self->{'outfile'}."prn", $self->{'outfile'};
		$word->Options->{PrintBackground} = $origback;
		}
	else {
		($self->{'outfile'} = $self->{'infile'}) =~ s/\.in$/.txt/;
		$doc->SaveAs( {
			'FileName' => $self->{'outfile'},
			'FileFormat' => $docoutform{$out_format}
			} );
		}
	$doc->close;
	}
		
sub xls_convert
	{
	my ($self, $out_format) = @_;
	my $newname;
	($newname = $self->{'infile'}) =~ s/\.in$/\.xls/;
	rename $self->{'infile'}, $newname;
	my $excel = ( Win32::OLE->GetActiveObject('Excel.Application')
		or Win32::OLE->new('Excel.Application', 'Quit') )
		or die Win32::OLE->LastError;
	my $wrk = $excel->Workbooks->Open($newname)
				or die Win32::OLE->LastError;
	if ($out_format eq 'xls5' or $out_format eq 'xls95') {
		($self->{'outfile'} = $self->{'infile'}) =~ s/\.in$/-out.xls/;
		$wrk->SaveAs( {
			'FileName' => $self->{'outfile'},
			'FileFormat' => $xlsoutform{$out_format}
			} );
		}
	else {
		($self->{'outfile'} = $self->{'infile'}) =~ s/\.in$/\.txt/;
		open FILEOUT, "> $self->{'outfile'}" or die "Error writing $self->{'outfile'}: $!";
		binmode FILEOUT;
		my $count = $wrk->Sheets->Count;
		my $plainfile;
		($plainfile = $self->{'infile'}) =~ s/\.in$//;
		for my $i (1 .. $count)
			{
			$wrk->Sheets($i)->Activate;
			#
			# Beware, this will append prn|csv to the FileName
			$wrk->SaveAs( {
				'FileName' => $plainfile,
				'FileFormat' => $xlsoutform{$out_format}
				} );
			printf FILEOUT $SHEET_SEP, $i if $i > 1;
			open FILEPART, $plainfile.'.'.$out_format;
			binmode FILEPART;
			while (<FILEPART>) { print FILEOUT $_; }
			close FILEPART;
			}
		close FILEOUT;
		unlink $plainfile.'.'.$out_format;
		}
	$wrk->{'Saved'} = 1;
	$wrk->close;
	}


sub result_length
	{
	my $self = shift;
	return -s $self->{'outfile'};
	}
sub get
	{
	my ($self, $len) = @_;
	my $fh = $self->{'outfh'};
	if (not defined $fh)
		{
		$fh = $self->{'outfh'} = new IO::File($self->{'outfile'});
		binmode $fh;
		}
	my $buffer;
	read $fh, $buffer, $len;
	$buffer;
	}
sub finished
	{
	my $self = shift;
	if (defined $self->{'fh'})
		{ delete $self->{'fh'}; }
	unlink $self->{'infile'} if defined $self->{'infile'};
	if (defined $self->{'outfh'})
		{ delete $self->{'outfh'}; }
	unlink $self->{'outfile'} if defined $self->{'outfile'};
	}
sub DESTROY
	{ shift->finished; }
1;

