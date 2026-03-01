class DocxError(Exception):
    """Base error for py-docx."""


class InvalidDocxError(DocxError):
    """Raised when a file is not a valid DOCX archive."""


class DocumentClosedError(DocxError):
    """Raised when operations are attempted after cleanup."""
