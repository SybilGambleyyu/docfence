"""Command-line interface for DocFence."""

from __future__ import annotations

import argparse
import os
import sys
import tempfile
from collections.abc import Sequence
from pathlib import Path

from docfence import __version__
from docfence.diff import diff_documents
from docfence.errors import DocFenceError, OutputError
from docfence.output import render_profile, render_report
from docfence.policy import apply_policy, load_policy, starter_policy
from docfence.snapshot import load_snapshot


def main(argv: Sequence[str] | None = None) -> int:
    """Run DocFence and return its documented process status."""

    parser = _parser()
    arguments = parser.parse_args(argv)
    try:
        if arguments.command == "profile":
            snapshot = load_snapshot(arguments.document)
            _write_output(render_profile(snapshot, arguments.format), arguments.output)
            return 0
        if arguments.command == "diff":
            report = diff_documents(arguments.before, arguments.after)
            _write_output(render_report(report, arguments.format), arguments.output)
            return 0
        if arguments.command == "check":
            report = diff_documents(arguments.before, arguments.after)
            report = apply_policy(report, load_policy(arguments.policy))
            _write_output(render_report(report, arguments.format), arguments.output)
            return 1 if report.findings else 0
        if arguments.command == "init":
            _write_starter_policy(arguments.path, arguments.force)
            return 0
    except DocFenceError as error:
        print(f"docfence: {error}", file=sys.stderr)
        return 2
    parser.error("a command is required")
    return 2


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="docfence",
        description="Local-first, privacy-safe DOCX/DOCM change assurance.",
    )
    parser.add_argument(
        "--version", action="version", version=f"%(prog)s {__version__}"
    )
    commands = parser.add_subparsers(dest="command", required=True)

    profile = commands.add_parser("profile", help="inventory one document")
    profile.add_argument("document")
    profile.add_argument("--format", choices=("json", "markdown"), default="json")
    profile.add_argument("--output", help="write atomically to this path")

    diff = commands.add_parser("diff", help="compare two documents")
    _comparison_arguments(diff, include_policy=False)

    check = commands.add_parser("check", help="compare documents and enforce a policy")
    _comparison_arguments(check, include_policy=True)

    init = commands.add_parser("init", help="write a conservative starter policy")
    init.add_argument("path", nargs="?", default="docfence.yml")
    init.add_argument(
        "--force", action="store_true", help="replace an existing regular file"
    )
    return parser


def _comparison_arguments(
    parser: argparse.ArgumentParser, *, include_policy: bool
) -> None:
    parser.add_argument("before")
    parser.add_argument("after")
    if include_policy:
        parser.add_argument("--policy", required=True)
    parser.add_argument(
        "--format", choices=("json", "markdown", "sarif"), default="json"
    )
    parser.add_argument("--output", help="write atomically to this path")


def _write_starter_policy(path: str, force: bool) -> None:
    try:
        target = Path(path)
        if target.exists() and not force:
            raise OutputError(
                "policy destination already exists; use --force to replace it"
            )
        if target.exists() and (target.is_symlink() or not target.is_file()):
            raise OutputError("policy destination must be a regular file")
    except (OSError, TypeError, ValueError):
        raise OutputError("policy destination cannot be inspected") from None
    _write_output(starter_policy(), path)


def _write_output(content: str, destination: str | None) -> None:
    if destination is None:
        sys.stdout.write(content)
        return

    temporary_path: Path | None = None
    try:
        target = Path(destination)
        if target.exists() and target.is_symlink():
            raise OutputError("output destination must not be a symbolic link")
        parent = target.parent
        if not parent.is_dir():
            raise OutputError("output directory does not exist")
        descriptor, temporary_name = tempfile.mkstemp(prefix=".docfence-", dir=parent)
        temporary_path = Path(temporary_name)
        with os.fdopen(descriptor, "w", encoding="utf-8", newline="\n") as stream:
            stream.write(content)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary_path, target)
        temporary_path = None
    except OutputError:
        raise
    except (OSError, TypeError, ValueError):
        raise OutputError("output cannot be written") from None
    finally:
        if temporary_path is not None:
            try:
                temporary_path.unlink(missing_ok=True)
            except OSError:
                pass


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
