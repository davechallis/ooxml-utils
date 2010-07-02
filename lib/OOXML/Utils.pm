package OOXML::Utils;
use strict;
use warnings;
use Carp;

BEGIN {
    use Exporter ();
    use vars qw($VERSION @ISA @EXPORT @EXPORT_OK %EXPORT_TAGS);
    $VERSION     = '0.01';
    @ISA         = qw(Exporter);
    #Give a hoot don't pollute, do not export more than needed by default
    @EXPORT      = qw();
    @EXPORT_OK   = qw();
    %EXPORT_TAGS = ();
}

sub get_namespace {
    my ($ns) = @_;
    my $base = 'http://schemas.openxmlformats.org';
    my $namespaces = {
'content_types' => "$base/package/2006/content-types",
'office_document' => "$base/officeDocument/2006/relationships/officeDocument",
'relationships' => "$base/package/2006/relationships",
'slide' => "$base/officeDocument/2006/relationships/slide",
'presentation' => "$base/presentationml/2006/main",
'document_relationships' => "$base/officeDocument/2006/relationships",
'image' => "$base/officeDocument/2006/relationships/image",
'drawing' => "$base/drawingml/2006/main",
    };

    if (!defined($namespaces->{$ns})) {
        carp("Namespace for $ns not defined");
    }
    return $namespaces->{$ns};
}

sub get_content_type {
    my ($ct) = @_;
    my $content_types = {
'slide' => 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
'relationship' => 'application/vnd.openxmlformats-package.relationships+xml',
    };

    if (!defined($content_types->{$ct})) {
        carp("Content-type for $ct not defined");
    }
    return $content_types->{$ct};
}

sub get_mime_type_from_filename {
    my ($filename) = @_;

    my @parts = split(/\./, $filename);
    my $ext = pop(@parts);

    my $mime_types = {
        'jpg'   => 'image/jpeg',
        'jpeg'  => 'image/jpeg',
        'gif'   => 'image/gif',
        'png'   => 'image/png'
    };

    if (!defined($mime_types->{$ext})) {
        carp("Mime-type for extension '$ext' not found");
    }
    return $mime_types->{$ext};
}

sub fatal {
    my ($msg, $exit_code) = @_;
    if (!defined($exit_code)) {
        $exit_code = 1;
    }

    print "[Error]: $msg\n";
    exit($exit_code);
}

sub warning {
   my ($msg) = @_;
   print "[Warning]: $msg\n";
}

#################### subroutine header begin ####################

=head2 sample_function

 Usage     : How to use this function/method
 Purpose   : What it does
 Returns   : What it returns
 Argument  : What it wants to know
 Throws    : Exceptions and other anomolies
 Comment   : This is a sample subroutine header.
           : It is polite to include more pod and fewer comments.

See Also   : 

=cut

#################### subroutine header end ####################


sub new
{
    my ($class, %parameters) = @_;

    my $self = bless ({}, ref ($class) || $class);

    return $self;
}


#################### main pod documentation begin ###################
## Below is the stub of documentation for your module. 
## You better edit it!


=head1 NAME

OOXML::Utils - Module abstract (<= 44 characters) goes here

=head1 SYNOPSIS

  use OOXML::Utils;
  blah blah blah


=head1 DESCRIPTION

Stub documentation for this module was created by ExtUtils::ModuleMaker.
It looks like the author of the extension was negligent enough
to leave the stub unedited.

Blah blah blah.


=head1 USAGE



=head1 BUGS



=head1 SUPPORT



=head1 AUTHOR

    Dave Challis
    CPAN ID: MODAUTHOR
    XYZ Corp.
    dsc@ecs.soton.ac.uk
    http://www.ecs.soton.ac.uk/people/dsc

=head1 COPYRIGHT

This program is free software licensed under the...

	The General Public License (GPL)
	Version 2, June 1991

The full text of the license can be found in the
LICENSE file included with this module.


=head1 SEE ALSO

perl(1).

=cut

#################### main pod documentation end ###################


1;
# The preceding line will help the module return a true value

