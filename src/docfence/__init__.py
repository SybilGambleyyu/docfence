"""DocFence: local-first DOCX change assurance."""

__version__ = "0.54.0"

from docfence.diff import diff_documents
from docfence.snapshot import load_snapshot

__all__ = ["__version__", "diff_documents", "load_snapshot"]
