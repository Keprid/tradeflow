"""
pipeline_errors.py
==================

Typed errors for the Trade Flow pipeline library.

Library functions (file finders, parsers, builders) raise PipelineError to
signal a user-facing failure.  The CLIs convert it to a clean sys.exit; the
webapp converts it to an HTTP 400 with the same message.  This replaces the
old pattern of calling sys.exit() deep inside importable library code, which
forced the webapp to catch SystemExit in ten different places.
"""


class PipelineError(Exception):
    """A user-facing pipeline failure (missing inputs, bad data, ...).

    The message is shown to the user as-is; keep it actionable.
    """
