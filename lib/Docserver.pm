
package Docserver;
use strict;
use Docserver::Config;
use IO::File;
use Fcntl;

BEGIN {
	local $^W = 0;
eval <<'EOF';
	use Win32;
	use Win32::API;
	use Win32::OLE;
	use Win32::OLE::Const 'Microsoft Office';
	use Win32::OLE::Const 'Microsoft Word';
	use Win32::OLE::Const 'Microsoft Excel';
EOF
}

$Docserver::VERSION = 0.981;

# Placeholder for global error string
use vars qw( $errstr );

# Values for output formats
my %docoutform = (
		'txt' =>	'Text with Layout',
		# 'txt1' =>	wdFormatTextLineBreaks,
		'txt1' =>	wdFormatText,
		'rtf' =>	wdFormatRTF,
		'doc' =>	wdFormatDocument,
		'doc6' =>	'MSWord6Exp',
		'doc95' =>	'MSWord6Exp',
		'html' =>	wdFormatHTML,
		'ps' =>		$Docserver::Config::Config{'ps'},
		'ps1' =>	$Docserver::Config::Config{'ps1'},
		); 
my %xlsoutform = (
		'prn' =>	xlTextPrinter,
		'txt' =>	xlTextPrinter,
		'csv' =>	xlCSV,
		'xls' =>	xlNormal,
		'xls5' =>	xlExcel5,
		'xls95' =>	xlExcel5,
		'html' =>	xlHtml,
		'ps' =>		'Adobe Generic PS on FILE:',
		'ps1' =>	'Adobe Generic PS1 on FILE:',
		);

my $SHEET_SEP = "\n##### List è. %s #####\n\n";
my $CSV_SHEET_SEP = "##### Sheet %s #####\n";

sub Debug {
	my $self = shift;
	$self->{'server'}->Debug(@_);
	}


# http://msdn.microsoft.com/library/psdk/psdkref/errlist_9usz.htm
# contains (rather long) table of system errors. You can also get
# there from http://msdn.microsoft.com/library/, should MS change the
# URL.
sub WinError ($$) {
	my ($self, $func) = @_;
	if ($^E) {
		$self->Debug("Win32 API function $func returned error: $^E");
		$^E = 0;
	}
}

# Create Docserver object, create temporary directory and create
# and open input (temporary storage) file in binmode.
sub new {
	my $class = shift;
	my $self;
	eval {
		$self = bless {
			'verbose' => 5,
			}, $class;

		my ($dir, $filename)
			= ($Docserver::Config::Config{'tmp_dir'}, 'file.in');
		if (not -d $dir) {
			print STDERR "Directory `$dir' doesn't exist, will try to create it\n";
			mkdir $dir, 0666
				or die "Error creating dir `$dir': $!\n";
			die "Directory `$dir' was not created properly\n" if not -d $dir;
		}
		$dir .= '\\'.time.'.'.$$;
		mkdir $dir, 0666 or die "Error creating tmp dir `$dir': $!\n";
		$self->{'dir'} = $dir;

		$self->{'infile'} = $dir.'\\'.$filename;
		print STDERR "Temporary file is `$self->{'infile'}'\n";

		$self->{'fh'} = new IO::File ">$self->{'infile'}"
			or die "Couldn't create file `$self->{'infile'}': $@\n";
		binmode $self->{'fh'};
	};
	if ($@) {
		$errstr = $@;
		return;
	}
	return $self;
}

# Returns error string, either for class or for object.
sub errstr {
	my $ref = shift;
	if (defined $ref and ref $ref) { return $ref->{'errstr'}; }
	return $errstr;
}

# Chooses smaller chunk size -- compares server configuration with
# value that came from client.
sub preferred_chunk_size {
	my ($self, $size) = @_;
	$size = $Docserver::Config::Config{'ChunkSize'}
		if not defined $size or $size > $Docserver::Config::Config{'ChunkSize'};
	print STDERR "Choosing chunk size `$size'\n" if $self->{'verbose'};
	$self->{'ChunkSize'} = $size;
	return $size;
}

# Sets the input file length in the object.
sub input_file_length {
	my ($self, $size) = @_;
	if (defined $size) {
		$self->{'input_file_length'} = shift;
		print STDERR "Setting input file size to `$size'\n" if $self->{'verbose'};
	}
	$size;
}

# Puts next chunk of data into the input file.
sub put {
	my $self = shift;
	print { $self->{'fh'} } shift;
	1;
}

# Runs the conversion from infile, from in_format to out_format.
sub convert {
	my ($self, $in_format, $out_format) = @_;
	delete $self->{'errstr'};
	print STDERR "Called convert (`$in_format', `$out_format')\n"
		if $self->{'verbose'};

	eval {
		if (defined $self->{'fh'}) {
			# Close the input filehandle, no more data coming.
			$self->{'fh'}->close();
			delete $self->{'fh'};
		}
		if ($in_format eq 'doc' or $in_format eq 'rtf') {
			# Run Word conversion.
			if (not defined $docoutform{$out_format}) {
				die "Unsupported output format `$out_format' for Word conversion\n";
			}
			$self->doc_convert($in_format, $out_format);
		}
		elsif ($in_format eq 'xls' or $in_format eq 'csv') {
			# Run Excel conversion.
			if (not defined $xlsoutform{$out_format}) {
				die "Unsupported output format `$out_format' for Excel conversion\n";
			}
			$self->xls_convert($in_format, $out_format);
		}
		else {
			die "Unsupported input format `$in_format'\n";
		}
	};
	if ($@) {
		$self->{'errstr'} = $@;
		print STDERR "Conversion failed: $@\n";
	}
	return 1 if not defined $self->{'errstr'};
	return;
}

# Does the whole conversion from Word doc.
sub doc_convert {
	my ($self, $in_format, $out_format) = @_;

	# We'll shift the file to *.doc.
	my $newname;
	if ($in_format eq 'rtf') {
		($newname = $self->{'infile'}) =~ s/\.[^\.]+$/.rtf/;
	} else {
		($newname = $self->{'infile'}) =~ s/\.[^\.]+$/.doc/;
	}
	rename $self->{'infile'}, $newname
		or die "Error moving `$self->{'infile'}' to `$newname': $!\n";
	$self->{'infile'} = $newname;

	print STDERR "Processing file `$newname'\n";

	# We will start new Word. It is better than doing
	# GetActiveObject because if the interactive user already has
	# some Word open, he won't see any documents flashing through
	# his screen, and vice verse, he shouldn't be able to spoil
	# our conversion. GetActiveObject would be necessary if we
	# wanted the user to be able to kill potential dialog windows.

	my $word = Win32::OLE->new('Word.Application', 'Quit')
		or die Win32::OLE->LastError;

	# Convertors for Text with Layout and to Word95 are optional
	# part of Word installation and as such they don't have any
	# constant wdFormat. They get some integer upon installation
	# and we need to get them this way.

	if (not $out_format =~ /^ps\d*$/
		and not $docoutform{$out_format} =~ /^\d+$/) {
		for my $i (1 .. $word->FileConverters->Count) {
			$docoutform{$out_format}
				= $word->FileConverters($i)->SaveFormat
					if ($word->FileConverters($i)->ClassName
						eq $docoutform{$out_format});
		}
		if (not $docoutform{$out_format} =~ /^\d+$/) {
			die "Couldn't find converter for format `$docoutform{$out_format}'\n";
		}
		print STDERR "Found output converter number `$docoutform{$out_format}'\n";
	}

	# Open the Word document. We have to distinguish here. It
	# would be nicer to call it the Open way (I'm not sure if
	# after adding new document it is always active), however Open
	# doesn't handle templates (unlike Excel), which cannot be
	# saved as anything else than templates. ConfirmConversions and
	# Format are here because interactive user could change these.

	my $doc;
	if ($in_format eq 'doc') {
		$word->Documents->Add({
			'Template' => $self->{'infile'},
			}) or die Win32::OLE->LastError;
		$doc = $word->ActiveDocument;
	} else {
		$doc = $word->Documents->Open({
			'FileName' => $self->{'infile'},
			'ConfirmConversions' => 0,
			'Format' => wdOpenFormatAuto,
			}) or die Win32::OLE->LastError;
	}

	if ($out_format =~ /^ps\d*$/) {
		# Print to .prn file. We have to run it on background
		# so that we don't get some dialog that would allow
		# interactive user to cancel the print.

		($self->{'outfile'} = $self->{'infile'}) =~ s/\.[^\.]+$/.prn/;
		my $origback = $word->Options->{PrintBackground};
		my $origprinter = $word->ActivePrinter;

		$word->Options->{PrintBackground} = 1;
		my $printer = $docoutform{$out_format};
		print STDERR "Setting ActivePrinter to `$printer'\n";
		$word->{ActivePrinter} = $printer;
		if ($word->{ActivePrinter} ne $printer) {
			print STDERR "ActivePrinter set to `$word->{ActivePrinter}'\n";
			die "Setting ActivePrinter to `$printer' failed -- printer not found\n";
		}

		$doc->Activate;
		$word->PrintOut({
			'Range' => wdPrintAllDocument,
			'PrintToFile' => 1, 
			'OutputFileName' => $self->{'outfile'},
			'Copies' => 1
			});
		for (my $i = 0; $i < 60; $i++) {
			sleep 2;
			last unless $word->{BackgroundPrintingStatus};
		}
		$word->Options->{PrintBackground} = $origback;
		$word->{ActivePrinter} = $origprinter;
	} else {
		# The Text with Layout has problems with pictures,
		# probably whenever the picture is in header or footer
		# (it issues error message that saving cannot be
		# finished because access rights are wrong), and
		# sometimes even in normal text when it produces
		# garbage.
		#
		# Because it seems to ignore shapes wich text fields
		# and probably of all types, we'll delete all shapes.
		# Because the count decreases when we delete the
		# shapes, we have to delete from the beginning and not
		# to walk throught the list. We have to delete shapes
		# from header and footer separately ($doc->Shapes
		# won't return them), according to the manual we can
		# take Shapes property from any HeaderFooter object
		# and the returned collecion will contain all shapes
		# from all headers and footers.
		#
		# Accessing $doc->Sections(1)->...->Count in for cycle
		# (even if it doesn't get executed) caused in trivial
		# test case with one line header and one line of
		# normal text an extra Enter to be inserted between
		# the header and the body. So we better not access
		# this Count.
		#
		# The normal text convertor puts header under the
		# normal text. Moreover, if the original document
		# containg page numbers in headers or footers, the
		# Text with Layout puts to the beginning of the output
		# (and the normal text converter to the end of the
		# output) some number, usually number of the first
		# page, but the following pages are not numbered.
		# That's why we'll remove all page numbers,
		# unfortunately there doesn't seem to be any better
		# way than walking through all combinations of header
		# and footer and constants WdHeaderFooterIndex.

		if ($out_format eq 'txt') {
			for my $shape (1 .. $doc->Shapes->Count) {
				$doc->Shapes(1)->Delete;
			}
			for my $shape (1 ..  $doc->Sections(1)->Headers(wdHeaderFooterPrimary)->Shapes->Count) {
				$doc->Sections(1)->Headers(wdHeaderFooterPrimary)->Shapes(1)->Delete;
			}
		}

		if ($out_format =~ /^txt1?$/) {
			for my $section (1 .. $doc->Sections->Count) {
				for my $pagenumger (1 .. $doc->Sections($section)->Footers(wdHeaderFooterPrimary)->PageNumbers->Count) {
					$doc->Sections($section)->Footers(wdHeaderFooterPrimary)->PageNumbers(1)->Delete;
				}
				for my $pagenumger (1 .. $doc->Sections($section)->Headers(wdHeaderFooterPrimary)->PageNumbers->Count) {
					$doc->Sections($section)->Headers(wdHeaderFooterPrimary)->PageNumbers(1)->Delete;
				}
				for my $pagenumger (1 .. $doc->Sections($section)->Footers(wdHeaderFooterEvenPages)->PageNumbers->Count) {
					$doc->Sections($section)->Footers(wdHeaderFooterEvenPages)->PageNumbers(1)->Delete;
				}
				for my $pagenumger (1 .. $doc->Sections($section)->Headers(wdHeaderFooterEvenPages)->PageNumbers->Count) {
					$doc->Sections($section)->Headers(wdHeaderFooterEvenPages)->PageNumbers(1)->Delete;
				}
				for my $pagenumger (1 .. $doc->Sections($section)->Footers(wdHeaderFooterFirstPage)->PageNumbers->Count) {
					$doc->Sections($section)->Footers(wdHeaderFooterFirstPage)->PageNumbers(1)->Delete;
				}
				for my $pagenumger (1 .. $doc->Sections($section)->Headers(wdHeaderFooterFirstPage)->PageNumbers->Count) {
					$doc->Sections($section)->Headers(wdHeaderFooterFirstPage)->PageNumbers(1)->Delete;
				}
			}
		}

		($self->{'outfile'} = $self->{'infile'}) =~ s/\.[^\.]+$/.out/;
		$doc->SaveAs({
			'FileName' => $self->{'outfile'},
			'FileFormat' => $docoutform{$out_format}
			});
	}
	$doc->Close({
		'SaveChanges' => wdDoNotSaveChanges,
		});
}
		
sub xls_convert
	{
	my ($self, $in_format, $out_format) = @_;
	my $newname;
	($newname = $self->{'infile'}) =~ s/\.[^\.]+$/\.xls/;
	rename $self->{'infile'}, $newname;

	my $excel = Win32::OLE->new('Excel.Application', 'Quit')
		or die Win32::OLE->LastError;

	my $wrk = $excel->Workbooks->Open({
		'FileName' => $self->{'infile'},
		($in_format eq 'csv' ? ('Format' => 4) : ()),
		}) or die Win32::OLE->LastError;

	# We'll set nice name of the sheet if the input comes from CSV
	# to have reasonable name for output to PS or XLS.
	$wrk->Sheets(1)->{'Name'} = 'Sheet1' if $in_format eq 'csv';

	($self->{'outfile'} = $self->{'infile'}) =~ s/\.[^\.]+$/.out/;
	if ($out_format =~ /^xls(95)?$/) {
		$wrk->SaveAs({
			'FileName' => $self->{'outfile'},
			'FileFormat' => $xlsoutform{$out_format}
			});

	} elsif ($out_format =~ /^ps\d?$/) {
		# It seems like Excel cannot do background printing
		# like Word can. Fortunately the dialog box is not
		# active so it cannot be hit by accident.
		#
		# We have to setup ActivePrinter and then return it
		# back because we could change the printer to
		# interactive Excel (but it's no longer a problem
		# because we start new Excel.Application each time.

		my $origprinter = $excel->ActivePrinter;
		$excel->{ActivePrinter} = $xlsoutform{$out_format};
		$wrk->Activate;
		$excel->PrintOut({
			'PrintToFile' => 1, 
			'PrToFileName' => $self->{'outfile'},
			});
		$excel->{ActivePrinter} = $origprinter;

	} elsif ($out_format eq 'csv') {
		my $savefile = $self->{'outfile'};
		$savefile =~ s/\.out$/.xout/;

		open FILEOUT, "> $self->{'outfile'}" or die "Error writing $self->{'outfile'}: $!";
		binmode FILEOUT;
		for my $i (1 .. $wrk->Sheets->Count) {
			$wrk->Sheets($i)->SaveAs({
				'FileName' => $savefile,
				'FileFormat' => $xlsoutform{$out_format},
			});
			printf FILEOUT $CSV_SHEET_SEP, $i if $i > 1;
			open IN, $savefile;
			binmode IN;
			while (<IN>) {
				print FILEOUT;
			}
			close IN;
			unlink $savefile;
		}
		close FILEOUT;
	
	} elsif ($out_format eq 'html') {
		my $savefile;
		for my $i (1 .. $wrk->Sheets->Count) {
			$savefile = $self->{'outfile'};
			$savefile =~ s/\.out$/$1.out/;
			$wrk->PublishObjects->Add({
				'SourceType' => xlSourceSheet,
				'Filename' => $savefile,
				'Sheet' => $wrk->Sheets($i)->Name,
				'HtmlType' => xlHtmlStatic,
				})->Publish({
					'Create' => 0,
					});

		}
	}

	$wrk->Close({
		'SaveChanges' => 0
		});
}


sub result_length {
	my $self = shift;
	return -s $self->{'outfile'};
}

# Returns next piece of output file.
sub get {
	my ($self, $len) = @_;
	my $fh = $self->{'outfh'};
	if (not defined $fh) {
		$fh = $self->{'outfh'} = new IO::File($self->{'outfile'});
		binmode $fh;
	}
	my $buffer;
	read $fh, $buffer, $len;
	$buffer;
}

sub finished {
	my $self = shift;

	### FIXME: vymazat soubory. Cely adresar.
return;
	if (defined $self->{'fh'})
		{ delete $self->{'fh'}; }
	unlink $self->{'infile'} if defined delete $self->{'infile'};
	if (defined $self->{'outfh'})
		{ delete $self->{'outfh'}; }
	unlink $self->{'outfile'} if defined delete $self->{'outfile'};
}

sub DESTROY {
	shift->finished;
}

sub server_version {
	return $Docserver::VERSION;
}

1;

=head1 NAME

Docserver.pm - server module for remote MS format conversions

=head1 AUTHOR

(c) 1998--2001 Jan Pazdziora, adelton@fi.muni.cz,
http://www.fi.muni.cz/~adelton/ at Faculty of Informatics, Masaryk
University in Brno, Czech Republic.

Pavel Smerk added support for more formats and also did the error
and Windows handling.

=cut

