package OOXML::Utils::Part;

use strict;
use warnings;

use XML::LibXML;
use OOXML::Utils;
use OOXML::Utils::RelationshipPart;

sub new {
    my ($class) = @_;
    my $self = {
        'relative_uri' => undef,
        'package' => undef,
        'uri_prefix' => undef,
        'absolute_uri' => undef
    };
    bless($self, $class);
    return $self;
}

sub get_relative_uri {
    my ($self) = @_;
    return $self->{'relative_uri'};
}

sub set_relative_uri {
    my ($self, $relative_uri) = @_;
    $self->{'relative_uri'} = $relative_uri;
}

sub get_absolute_uri {
    my ($self) = @_;
    return $self->{'absolute_uri'};
}

sub set_absolute_uri {
    my ($self, $absolute_uri) = @_;
    $self->{'absolute_uri'} = $absolute_uri;
}

sub get_package {
    my ($self) = @_;
    return $self->{'package'};
}

sub set_package {
    my ($self, $package) = @_;
    $self->{'package'} = $package;
}

sub get_uri_prefix {
    my ($self) = @_;
    return $self->{'uri_prefix'};
}

sub set_uri_prefix {
    my ($self, $uri_prefix) = @_;
    $self->{'uri_prefix'} = $uri_prefix;
}

sub get_uri {
    my ($self) = @_;
    
    # Use absolute if defined
    if (defined($self->get_absolute_uri())) {
        return $self->get_absolute_uri();
    }

    # Else build URI from relative and prefix
    if (defined($self->get_uri_prefix())) {
        return $self->get_uri_prefix() . '/' . $self->get_relative_uri();
    }

    return $self->get_relative_uri();
}

sub set_uri {
    my ($self, $uri) = @_;
    $self->set_absolute_uri($uri);
}

sub get_content_type {
    my ($self) = @_;

    my $ct = $self->get_package()->get_part('[Content_Types].xml');
    my $doc = $ct->get_xmldoc();
    my $ct_ns = OOXML::Utils::get_namespace('content_types');
    my $part_uri = $self->get_uri();

    foreach my $override ($doc->getElementsByTagNameNS($ct_ns, 'Override')) {
        if ($override->hasAttribute('PartName')) {
            if ($override->getAttribute('PartName') eq "/$part_uri") {
                return $override->getAttribute('ContentType');
            }
        }
    }

    return 0;
}

sub get_relationships_uri {
    my ($self) = @_;
    my @uri_parts = split('/', $self->get_uri());
    my $part_name = pop(@uri_parts);
    my $base_uri = join('/', @uri_parts);
    return "$base_uri/_rels/$part_name.rels";
}

sub get_relationships_part {
    my ($self) = @_;
    my $rp = $self->get_package()->get_part($self->get_relationships_uri());
    $rp->set_uri_prefix($self->get_uri_prefix());
    return $rp;
}

# Get array of XML elements
sub get_relationships_as_nodelist {
    my ($self) = @_;

    my $rel_ns = OOXML::Utils::get_namespace('relationships');

    my $part = $self->get_relationships_part();
    my $doc = $part->get_xmldoc();

    return $doc->getElementsByTagNameNS($rel_ns, 'Relationship');
}


sub get_xmldoc {
    my ($self) = @_;
    return $self->get_package()->get_xmldoc_from_part_uri($self->get_uri());
}

# replace part in zip from xmldoc
sub set_xmldoc {
    my ($self, $xmldoc) = @_;
    $self->set_contents($xmldoc->toString());
}

sub set_contents {
    my ($self, $contents) = @_;
    $self->get_package()->set_contents_for_part_uri($self->get_uri(),
                                                    $contents);
}

sub get_new_relationship_id {
    my ($self) = @_;

    # Start at 1 and find next unused ID
    my $rel_id = 1;
    foreach my $rel ($self->get_relationships_as_nodelist()) {
        if ($rel->hasAttribute('Id')) {
            my $id = $rel->getAttribute('Id');
            if ($id =~ m/^rId([0-9]+)$/ and $1 >= $rel_id) {
                $rel_id = $1 + 1;
            }
        }
    }

    return $rel_id;
}

sub delete {
    my ($self) = @_;
    return $self->get_package()->remove_member($self->get_uri());
}

1;
