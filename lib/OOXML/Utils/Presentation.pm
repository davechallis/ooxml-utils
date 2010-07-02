package OOXML::Utils::Presentation;

use strict;
use warnings;
use OOXML::Utils;
use OOXML::Utils::Exceptions qw( :document_exceptions );

use base qw( OOXML::Utils::Document );

# Get count of number of slides in this presentation
sub get_number_of_slides {
    my ($self) = @_;

    my $part_uri = $self->get_package()->get_main_part();
    my $slide_rel = OOXML::Utils::get_namespace('slide');

    my $n_slides = 0;
    my @relationships =
        $self->get_package()->get_relationships_from_part_uri($part_uri);

    foreach my $relationship (@relationships) {
        if ($relationship->hasAttribute('Type')) {
            if ($relationship->getAttribute('Type') eq $slide_rel) {
                $n_slides += 1;
            }
        }
    }

    return $n_slides;
}

sub get_slide_rel_id_by_number {
    my ($self, $number) = @_;

    my $pres_ns = OOXML::Utils::get_namespace('presentation');
    my $doc_rel_ns =
        OOXML::Utils::get_namespace('document_relationships');

    # Ger presentation part
    my $pres_part = $self->get_package()->get_part('ppt/presentation.xml');
    $pres_part->set_uri_prefix('ppt');

    my $pres_xmldoc = $pres_part->get_xmldoc();
    my $slide_list = 
        pop(@{$pres_xmldoc->getElementsByTagNameNS($pres_ns, 'sldIdLst')});

    # Get slide from slide number and check it exists
    my $n = 1;
    my $src_slide_elem = undef;
    foreach my $elem ($slide_list->getElementsByTagNameNS($pres_ns, 'sldId')) {
        if ($n == $number) {
            $src_slide_elem = $elem;
            last;
        }
        $n += 1;
    }

    # Throw an exception if slide number asked for is out of range
    if (!defined($src_slide_elem)) {
        my $msg = "Slide number $number does not exist.";
        e_not_exists('message'=>$msg);
    }

    return $src_slide_elem->getAttributeNS($doc_rel_ns, 'id');
}

# slides/slide4.xml
sub get_new_slide_uri {
    my ($self) = @_;

    my $ct_ns = OOXML::Utils::get_namespace('content_types');
    my $slide_ct = OOXML::Utils::get_content_type('slide');

    my $ct_xmldoc =
        $self->get_package()->get_part('[Content_Types].xml')->get_xmldoc();
    my $slide_id = 1;

    my @elements = $ct_xmldoc->getElementsByTagNameNS($ct_ns, 'Override');
    foreach my $elem (@elements) {
        if ($elem->hasAttribute('ContentType') and
            $elem->getAttribute('ContentType') eq $slide_ct) {
            my $partname = $elem->getAttribute('PartName');
            if ($partname =~ m/slide([0-9]+)\.xml$/) {
                if ($1 >= $slide_id) {
                    $slide_id = $1 + 1;
                }
            }
        }
    }

    return 'slides/slide' . $slide_id . '.xml';
}

sub append_slide_from_existing_slide {
    my ($self, $src_slide_number) = @_;

    # Set up namespaces and content types needed
    my $slide_ct = OOXML::Utils::get_content_type('slide');
    my $slide_ns = OOXML::Utils::get_namespace('slide');
    my $ct_ns = OOXML::Utils::get_namespace('content_types');
    my $rel_ns = OOXML::Utils::get_namespace('relationships');
    my $rel_ct = OOXML::Utils::get_content_type('relationship');
    my $pres_ns = OOXML::Utils::get_namespace('presentation');

    # Get slide rel ID and confirm that it exists
    my $src_slide_rel_id =
        $self->get_slide_rel_id_by_number($src_slide_number);

    # Find an unused slide id in [Content_Types].xml
    my $slide_uri = $self->get_new_slide_uri();

    # Add new parts to [Content_Types].xml
    my $ct_part = $self->get_package()->get_part('[Content_Types].xml');
    my $ct_xmldoc = $ct_part->get_xmldoc();

    # Add slide part
    my $node = XML::LibXML::Element->new('Override');
    $node->setAttribute('PartName', "/ppt/$slide_uri");
    $node->setAttribute('ContentType', $slide_ct);
    $ct_xmldoc->getDocumentElement()->addChild($node);

    # Add slide relation
    $node = XML::LibXML::Element->new('Override');
    my $dest_slide_part = OOXML::Utils::Part->new();
    $dest_slide_part->set_package($self->get_package());
    $dest_slide_part->set_uri_prefix('ppt');
    $dest_slide_part->set_relative_uri($slide_uri);
    #$node->setAttribute(
    #    'PartName',
    #    '/' . $dest_slide_part->get_relationships_uri()
    #);
    #$node->setAttribute('ContentType', $rel_ct);
    #$ct_xmldoc->getDocumentElement()->addChild($node);

    # Write changes to [Content_Types].xml
    $ct_part->set_xmldoc($ct_xmldoc);

    # Ger presentation part
    my $pres_part = $self->get_package()->get_part('ppt/presentation.xml');
    $pres_part->set_uri_prefix('ppt');

    # Add new relationship to slide
    my $rel_id = $pres_part->get_new_relationship_id();

    my $rel_node = XML::LibXML::Element->new('Relationship');
    $rel_node->setAttribute('Id', 'rId' . $rel_id);
    $rel_node->setAttribute('Type', $slide_ns);
    $rel_node->setAttribute('Target', $slide_uri);

    my $pres_rel_part = $pres_part->get_relationships_part();
    my $rel_xmldoc = $pres_rel_part->get_xmldoc();
    my $rel_elem =
        pop(@{$rel_xmldoc->getElementsByTagNameNS($rel_ns, 'Relationships')});

    $rel_elem->addChild($rel_node);

    # Write changes to presentation.xml
    $pres_rel_part->set_xmldoc($rel_xmldoc);

    # Add slide to presentation.xml
    my $pres_xmldoc = $pres_part->get_xmldoc();
    my $slide_list = 
        pop(@{$pres_xmldoc->getElementsByTagNameNS($pres_ns, 'sldIdLst')});

    my $max_sid = 0;
    foreach my $elem ($slide_list->getElementsByTagNameNS($pres_ns, 'sldId')) {
        my $id = $elem->getAttribute('id');
        if ($id > $max_sid) {
            $max_sid = $id;
        }
    }

    my $sld_node = XML::LibXML::Element->new('p:sldId');
    $sld_node->setAttribute('id', $max_sid + 1);
    $sld_node->setAttribute('r:id', 'rId' . $rel_id);
    $slide_list->addChild($sld_node);
    $pres_part->set_xmldoc($pres_xmldoc);

    
    # Copy source slide to destination slide
    my $src_slide_part =
        $pres_rel_part->get_relationship_part_by_id($src_slide_rel_id);
    my $src_slide_xmldoc = $src_slide_part->get_xmldoc();

    $dest_slide_part->set_xmldoc($src_slide_xmldoc);

    # Copy source relationships to destination relationships
    my $dest_slide_rel_part = $dest_slide_part->get_relationships_part();
    $dest_slide_rel_part->set_xmldoc($src_slide_part->get_relationships_part()->get_xmldoc());
}

sub delete_slide {
    my ($self, $slide_num) = @_;

    my $slide_part = $self->get_slide_part_by_number($slide_num);
    my $slide_rel_part = $slide_part->get_relationships_part();

    my $ct_part = $self->get_package()->get_part('[Content_Types].xml');
    my $ct_xmldoc = $ct_part->get_xmldoc();
    
    my $xpc = XML::LibXML::XPathContext->new();
    $xpc->registerNs('ct', OOXML::Utils::get_namespace('content_types'));
    $xpc->registerNs('p', OOXML::Utils::get_namespace('presentation'));
    $xpc->registerNs('r', OOXML::Utils::get_namespace('document_relationships'));
    my $xpath_pre = '/ct:Types/ct:Override[@PartName=\'';
    my $xpath_post = '\']';

    # Remove slide from [Content_Types].xml
    my $xpath = $xpath_pre . '/' . $slide_part->get_uri() . $xpath_post;
    foreach my $node ($xpc->findnodes($xpath, $ct_xmldoc)) {
        $node->parentNode()->removeChild($node);        
    }

    # Remove slide relation from [Content_Types].xml
    $xpath = $xpath_pre . '/' . $slide_rel_part->get_uri() . $xpath_post;
    foreach my $node ($xpc->findnodes($xpath, $ct_xmldoc)) {
        $node->parentNode()->removeChild($node);        
    }

    # Write [Content_Types].xml changes
    $ct_part->set_xmldoc($ct_xmldoc);

    # Remove slide from presentation.xml
    my $slide_rid = $self->get_slide_rel_id_by_number($slide_num);
    my $pres_part = $self->get_package()->get_part('ppt/presentation.xml');
    $pres_part->set_uri_prefix('ppt');

    my $pres_xmldoc = $pres_part->get_xmldoc();
    $xpath = "/p:presentation/p:sldIdLst/p:sldId[\@r:id='$slide_rid']";
    foreach my $node ($xpc->findnodes($xpath, $pres_xmldoc)) {
        $node->parentNode()->removeChild($node);
    }
    $pres_part->set_xmldoc($pres_xmldoc);

    # Remove slide contents
    $slide_part->delete();

    # Remove slide relations
    $slide_rel_part->delete();
}

sub get_slide_part_by_number {
    my ($self, $slide_num) = @_;

    # Get rId number for slide n
    my $slide_rel_id = $self->get_slide_rel_id_by_number($slide_num);

    # Get reference to main presentation
    my $pres_part = $self->get_package()->get_part('ppt/presentation.xml');
    $pres_part->set_uri_prefix('ppt');

    # Get relations from main presentation, pull out rel with rId above
    my $pres_rel_part = $pres_part->get_relationships_part();
    return $pres_rel_part->get_relationship_part_by_id($slide_rel_id);
}

sub replace_slide_text {
    my ($self, $slide_num, $source_text, $replacement_text) = @_;

    my $slide_part = $self->get_slide_part_by_number($slide_num);
    my $xml = $slide_part->get_xmldoc();

    my $n_replacements = 0;
    
    # Find all text nodes, and replace source with dest text
    foreach my $node ($xml->findnodes('//text()')) {
        my $text = $node->nodeValue();
        my $source = quotemeta($source_text);
        if ($text =~ m/$source/) {
            $text =~ s/$source/$replacement_text/g;
            $node->setData($text);
            $n_replacements += 1;
        }
    }

    if ($n_replacements > 0) {
        $slide_part->set_xmldoc($xml);
    }

    return $n_replacements;
}

sub resize_slide_image {
    my ($self, $slide_num, $image_num, $width, $height, $units) = @_;

    my $p_ns = OOXML::Utils::get_namespace('presentation');
    my $a_ns = OOXML::Utils::get_namespace('drawing');

    my $cx = undef;
    my $cy = undef;

    # 1 pixel is roughly 9525 EMUs (EMUs are absolute units that can be
    # converted to inches/centimeters, size relative to screen DPI)
    if ($units eq 'pixels') {
        $cx = $width * 9525;
        $cy = $height * 9525;
    }

    my $slide_part = $self->get_slide_part_by_number($slide_num);
    my $slide_xml = $slide_part->get_xmldoc();

    # Loop until we find the image number we're after, then replace sizes
    my $m = 1;
    foreach my $elem ($slide_xml->getElementsByTagNameNS($p_ns, 'pic')) {
        if ($m == $image_num) {
            my $ext = $elem->getElementsByTagNameNS($a_ns, 'ext')->item(0);
            $ext->setAttribute('cx', $cx);
            $ext->setAttribute('cy', $cy);
            $slide_part->set_xmldoc($slide_xml);
            last;
        }
        $m += 1;
    }
}

sub replace_slide_image {
    my ($self, $slide_num, $image_num, $image_path) = @_;

    my $draw_ns = OOXML::Utils::get_namespace('drawing');
    my $docrel_ns = OOXML::Utils::get_namespace('document_relationships');
    my $img_ns = OOXML::Utils::get_namespace('image');
    my $rel_ns = OOXML::Utils::get_namespace('relationships');

    my $slide_part = $self->get_slide_part_by_number($slide_num);
    my $slide_rel_part = $slide_part->get_relationships_part();

    my @fparts = split(/\./, $image_path);
    my $image_extension = pop(@fparts);
    my $mime_type = OOXML::Utils::get_mime_type_from_filename($image_path);

    # Check mime type is defined in [Content_Types].xml
    my $ct_part = $self->get_package()->get_part('[Content_Types].xml');
    my $ct_xml = $ct_part->get_xmldoc();

    my $found = 0;
    my $last_elem = undef;
    foreach my $elem ($ct_xml->getElementsByTagName('Default')) {
        $last_elem = $elem;
        if ($elem->hasAttribute('Extension') &&
            $elem->getAttribute('Extension') eq $image_extension) {
            if ($elem->hasAttribute('ContentType') &&
                $elem->getAttribute('ContentType') eq $mime_type) {
                $found = 1; 
                last;
            }
        }
    }


    # If no content type for this extension found, add it
    if (!$found) {
        my $elem = $ct_xml->createElement('Default');
        $elem->setAttribute('Extension', $image_extension);
        $elem->setAttribute('ContentType', $mime_type);
        #$last_elem->addSibling($elem);

        # Add new elem to <Types>
        if (defined($last_elem)) {
            $ct_xml->firstChild->insertAfter($elem, $last_elem);
        }
        else {
            my @types = $ct_xml->getElementsByTagName('Types');
            my $types_elem = pop(@types);
            $types_elem->addChild($elem);
            #$ct_xml->addChild($elem);
        }
        $ct_part->set_xmldoc($ct_xml);
    }

    # Copy image file into media directory
    # Find an unused image name
    my $base_img_path = 'ppt/media/image';
    my $img_uri = undef;
    my $n = 1;
    while (!defined($img_uri)) {
        my $uri = "$base_img_path$n.$image_extension";
        if (!$self->get_package()->has_part($uri)) {
            $img_uri = $uri;
        }
        $n += 1;
    }

    # Add the image file to the zip
    $self->get_package()->set_part_from_file($img_uri, $image_path);

    # Work out which image to replace, and get it's relationship ID
    my $slide_xml = $slide_part->get_xmldoc();

    my $m = 1;
    my $rid = undef;
    foreach my $elem ($slide_xml->getElementsByTagNameNS($draw_ns, 'blip')) {
        if ($m == $image_num) {
            $rid = $elem->getAttributeNS($docrel_ns, 'embed');
            last;
        }
        $m += 1;
    }

    if (!defined($rid)) {
        my $msg = "Image number $image_num does not exist.";
        e_not_exists('message'=>$msg);
    }

    # Set relationships for slide to the new image
    my $sld_rel_xml = $slide_rel_part->get_xmldoc();
    my @nodes = $sld_rel_xml->getElementsByTagNameNS($rel_ns, 'Relationship');
    foreach my $elem (@nodes) {
        if ($elem->getAttribute('Id') eq $rid) {
            # Target needs to be relative uri
            my @img_parts = split(/\//, $img_uri);
            my $uri = '../media/' . pop(@img_parts);
            $elem->setAttribute('Target', $uri);
        }
    }

    $slide_rel_part->set_xmldoc($sld_rel_xml);
}

sub delete_slide_image {
    my ($self, $slide_num, $image_num) = @_;
}

# add_slide_from_slide_master
# move_slide_to_position
# add_slide_from_presentation


1;
