package OOXML::Utils::Document;

use strict;
use warnings;

use OOXML::Utils::Package;

sub new {
    my ($class) = @_;
    my $self = {'package'   => undef,
                'type'      => undef};
    bless($self, $class);
    return $self;
}

sub write_to_file {
    my ($self, $path) = @_;
    if (!defined($self->{'package'})) {
        return 0;
    }

    $self->{'package'}->write_to_file($path);
    return 1;
}

sub init_from_file {
    my ($self, $path) = @_;
    $self->{'package'} = OOXML::Utils::Package->new();
    return $self->{'package'}->init_from_file($path);
}

sub get_package {
    my ($self) = @_;
    return $self->{'package'};
}

# Wordprocessing, Spreadsheet or Presentation
sub get_type {
    my ($self) = @_;

    if (defined($self->{'type'})) {
        return $self->{'type'};
    }
}

1;
