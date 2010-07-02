package OOXML::Utils::Package;

use strict;
use warnings;

use Archive::Zip qw( :ERROR_CODES :CONSTANTS );
use File::Temp;
use File::Spec;
use File::Basename qw( fileparse );
use File::Copy qw( copy );
use XML::LibXML;
use OOXML::Utils;
use OOXML::Utils::Part;
use OOXML::Utils::Exceptions qw( :file_exceptions );

sub new {
    my ($class) = @_;
    my $self = {'zip' => undef};
    bless($self, $class);
    return $self;
}

sub get_zip {
    my ($self) = @_;
    return $self->{'zip'};
}

sub set_zip {
    my ($self, $zip) = @_;
    $self->{'zip'} = $zip;
}

# Write the zip file this uses to a given path
sub write_to_file {
    my ($self, $path) = @_;
    if ($self->get_zip()->writeToFileNamed($path) != AZ_OK) {
        e_file_not_writable('path'=>$path);
    }
}

# Create a temporary copy of the file, modify that as needed
sub init_from_file {
    my ($self, $path) = @_;

    if (!-e $path) {
        e_file_not_found('path'=>$path);
    }

    if (!-r $path) {
        e_file_not_readable('path' => $path);
    }

    $self->{'original_file_path'} = $path;

    # Create a temp file based on original file name
    my $original_filename = fileparse($path);
    my $temp_file = 
        File::Temp->new('TEMPLATE' => "$original_filename.XXXXX",
                        'UNLINK'   => 0,
                        'DIR'      => File::Spec->tmpdir());

    # Copy the original file to the temporary file
    my $temp_path = $temp_file->filename();
    $self->{'temp_file_path'} = $temp_path;
    copy($path, $temp_path)
        or e_file_not_writable('path'=>$temp_path);

    my $zip = Archive::Zip->new();
    if ($zip->read($temp_path) != AZ_OK) {
        e_file_not_readable('path'=>$temp_path);
    }

    $self->set_zip($zip);
}

sub DESTROY {
    my ($self) = @_;
    if (defined($self->{'temp_file_path'})) {
        if (-e $self->{'temp_file_path'}) {
            unlink($self->{'temp_file_path'});
        }
    }
}

sub is_part_name_valid {
    my ($self, $name) = @_;
    my @non_escaped = qw/ ! $ % & ' ( ) * + - . : ; = @ * _ ~ /;
    push(@non_escaped, ','); # , is added separately to avoid warning
    @non_escaped = map(quotemeta, @non_escaped);

    # TODO: check for hex codes better
    my $valid = '^(?:'
              . '[a-zA-Z0-9]|%[0-9a-fA-F]{2}|'
              . join('|', @non_escaped)
              . ')+$';
    if ($name =~ m/$valid/) {
        return 1;
    }
    return 0;
}

sub get_members {
    my ($self) = @_;
    return $self->get_zip()->members();
}

sub remove_member {
    my ($self, $uri) = @_;
    return $self->get_zip()->removeMember($uri);
}

sub get_xmldoc_from_part_uri {
    my ($self, $part_uri) = @_;

    my $member = $self->get_zip()->memberNamed($part_uri);
    my $parser = XML::LibXML->new();
    return $parser->parse_string($member->contents());
}

sub get_main_part {
    my ($self) = @_;
    my $od_ns = OOXML::Utils::get_namespace('office_document');
    my $rel_ns = OOXML::Utils::get_namespace('relationships');

    my $doc = $self->get_xmldoc_from_part_uri('_rels/.rels');

    my @relationships = $doc->getElementsByTagNameNS($rel_ns, 'Relationship');
    foreach my $relationship (@relationships) {
        if ($relationship->hasAttribute('Type')) {
            if ($relationship->getAttribute('Type') eq $od_ns) {
                return $relationship->getAttribute('Target');
            }
        }
    }

    return 0;
}

sub get_relationships_from_part_uri {
    my ($self, $part_uri) = @_;

    my @uri_parts = split('/', $part_uri);
    my $part_name = pop(@uri_parts);
    my $base_uri = join('/', @uri_parts);
    my $part_rel_uri = "$base_uri/_rels/$part_name.rels";

    my $xmldoc = $self->get_xmldoc_from_part_uri($part_rel_uri);

    my $rel_ns = OOXML::Utils::get_namespace('relationships');
    return $xmldoc->getElementsByTagNameNS($rel_ns, 'Relationship');
}

sub get_content_type_for_part {
    my ($self, $part_uri) = @_;

    my $doc = $self->get_xmldoc_from_part_uri('[Content_Types].xml');
    my $ct_ns = OOXML::Utils::get_namespace('content_types');

    foreach my $override ($doc->getElementsByTagNameNS($ct_ns, 'Override')) {
        if ($override->hasAttribute('PartName')) {
            if ($override->getAttribute('PartName') eq "/$part_uri") {
                return $override->getAttribute('ContentType');
            }
        }
    }

    return 0;
}

sub get_part {
    my ($self, $uri) = @_;

    my $part = undef;
    if ($uri =~ m/[.]rels$/) {
        $part = OOXML::Utils::RelationshipPart->new();
    }
    else {
        $part = OOXML::Utils::Part->new();
    }

    $part->set_uri($uri);
    $part->set_package($self);

    return $part;
}

sub has_part {
    my ($self, $uri) = @_;

    if (defined($self->get_zip()->memberNamed($uri))) {
        return 1;
    }
    return 0;
}

sub set_contents_for_part_uri {
    my ($self, $uri, $contents) = @_;

    if ($self->has_part($uri)) {
        $self->get_zip()->contents($uri, $contents);
    }
    else {
        $self->get_zip()->addString($contents, $uri);
    }
}

sub set_part_from_file {
    my ($self, $uri, $file_path) = @_;
    if ($self->has_part($uri)) {
        $self->get_zip()->updateMember($uri, $file_path);
    }
    else {
        $self->get_zip()->addFile($file_path, $uri);
    }
}

1;
