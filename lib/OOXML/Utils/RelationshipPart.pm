package OOXML::Utils::RelationshipPart;

use strict;
use warnings;

use base qw( OOXML::Utils::Part );

# Get array of XML elements
sub get_relationships_as_nodelist {
    my ($self) = @_;

    my $rel_ns = OOXML::Utils::get_namespace('relationships');

    return $self->get_xmldoc()->getElementsByTagNameNS($rel_ns, 'Relationship');
}

sub get_relationship_target_by_id {
    my ($self, $rid) = @_;

    my $elem = $self->get_relationship_element_by_id($rid);
    if ($elem) {
        return $elem->getAttribute('Target');
    }

    return 0;
}

sub get_relationship_element_by_id {
    my ($self, $rid) = @_;

    foreach my $elem ($self->get_relationships_as_nodelist()) {
        if ($elem->getAttribute('Id') eq $rid) {
            return $elem;
        }
    }

    return 0;
}

sub get_relationship_part_by_id {
    my ($self, $rid) = @_;

    my $target = $self->get_relationship_target_by_id($rid);
    my $uri = $self->get_uri_prefix() . '/' . $target;
    return $self->get_package()->get_part($uri);
}

1;
