package OOXML::Utils::Exceptions;

use base 'Exporter';
use Exception::Class;

use Exception::Class(
    'OOXML::Utils::Exceptions::FileException' => {
        'fields' => 'path',
        'description' => 'File and filesystem related exceptions'
    },
    'OOXML::Utils::Exceptions::FileNotFoundException' => {
        'isa' => 'OOXML::Utils::Exceptions::FileException',
        'alias' => 'e_file_not_found'
    },
    'OOXML::Utils::Exceptions::FileNotReadableException' => {
        'isa' => 'OOXML::Utils::Exceptions::FileException',
        'alias' => 'e_file_not_readable'
    },
    'OOXML::Utils::Exceptions::FileNotWritableException' => {
        'isa' => 'OOXML::Utils::Exceptions::FileException',
        'alias' => 'e_file_not_writable'
    },
    'OOXML::Utils::Exceptions::DocumentException' => {
        'fields' => 'message',
        'description' => 'Errors related to OOXML documents'
    },
    'OOXML::Utils::Exceptions::NotExists' => {
        'isa' => 'OOXML::Utils::Exceptions::DocumentException',
        'alias' => 'e_not_exists'
    },
);

@OOXML::Utils::Exceptions::EXPORT_OK = (
    'e_file_not_found',
    'e_file_not_readable',
    'e_file_not_writable',
    'e_not_exists'
);

%OOXML::Utils::Exceptions::EXPORT_TAGS = (
    'file_exceptions' => ['e_file_not_found',
                          'e_file_not_readable',
                          'e_file_not_writable'],
    'document_exceptions' => ['e_not_exists'],
);

1;
