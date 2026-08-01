"""Stable, non-content-bearing error types for DocFence."""


class DocFenceError(Exception):
    """Base class for expected DocFence failures."""


class DocumentFormatError(DocFenceError):
    """The input is not a supported, inspectable Word OOXML package."""


class DocumentSafetyError(DocFenceError):
    """The input crossed a package or XML safety boundary."""


class PolicyError(DocFenceError):
    """A policy could not be read or did not match the strict schema."""


class OutputError(DocFenceError):
    """A requested report destination cannot be safely written."""
